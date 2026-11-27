"""Temporary validation: actual delta traversal, synthetic 100K item corpus."""
import json
import statistics
import time
from collections import Counter
from uuid import UUID
from estimators.file_estimator import FileEstimator
from util.utils import ScanConfig
from util.thread_safe_ds import ThreadSafeMap

N = 100_000
PAGE = 1000
CREATORS = 1000
drive = 'drive'
users = [str(UUID(int=i+1)) for i in range(CREATORS)]
root = {'id':'root', 'folder':{}, 'parentReference':{'driveId':drive}}
plain = [root] + [
    {'id':str(i), 'file':{}, 'size':1024, 'parentReference':{'driveId':drive,'id':'root'}}
    for i in range(N)]
enriched = [root] + [dict(item, createdBy={'user':{'id':users[i % CREATORS]}})
                     for i, item in enumerate(plain[1:])]
def pages(items):
    return [json.dumps({'value':items[i:i+PAGE], **(
        {'@odata.nextLink':f'https://graph.microsoft.com/v1.0/page/{i//PAGE+1}'}
        if i+PAGE < len(items) else {'@odata.deltaLink':'done'})})
        for i in range(0,len(items),PAGE)]
payloads = {False:pages(plain),True:pages(enriched)}

def run(enabled):
    calls = []
    class Invoker:
        def invoke(self, base, batch, *args):
            responses = []
            for req in batch:
                url = req['url']
                index = int(url.rsplit('/', 1)[-1]) if '/page/' in '/' + url else 0
                calls.append(url)
                responses.append({'id':req['id'], 'status':200,
                                  'body':json.loads(payloads[enabled][index])})
            return responses
    config = ScanConfig(tenant_id='t',client_ids=['id'],client_secrets=['s'],
                        user_source='tenant',csv_path='',concurrency=2)
    e = FileEstimator(config,Invoker(),logger=lambda *a:None,
                      progress_update_callback=lambda *a,**k:None)
    p = ThreadSafeMap()
    for k in ('folderCount','fileCount','shortcutCount','versionCount'):
        p.update(k,0)
    counts = {drive:Counter()} if enabled else None
    completed = set() if enabled else None
    failures=[]
    start=time.perf_counter()
    try:
        result=e._create_in_memory_tree([drive],p,failures,
            creator_counts_by_drive=counts,creator_completed_drives=completed)
        elapsed=time.perf_counter()-start
    finally:
        e.shutdown()
    assert not failures, failures
    assert p.get_all()['fileCount']==N, p.get_all()
    assert len(calls)==len(payloads[enabled])
    if enabled:
        assert completed=={drive}
        assert len(counts[drive])==CREATORS
        assert sum(counts[drive].values())==N
    return elapsed, len(calls), result[3:]
if __name__=='__main__':
    for mode in (False,True):
        run(mode)
    times={False:[],True:[]}
    for i in range(5):
        results={}
        for mode in (False,True) if i%2==0 else (True,False):
            t,c,m=run(mode)
            times[mode].append(t)
            results[mode]=(c,m)
        assert results[False]==results[True]
    base,new=(statistics.median(times[m]) for m in (False,True))
    print(json.dumps({'items':N,'creators':CREATORS,'requests_each':len(payloads[False]),
       'baseline_seconds':base,'enabled_seconds':new,'cpu_and_json_overhead_pct':(new/base-1)*100,
       'baseline_payload_bytes':sum(map(len,payloads[False])),
       'enabled_payload_bytes':sum(map(len,payloads[True])),
       'baseline_samples':times[False],'enabled_samples':times[True]},indent=2))
