import json, urllib.request, urllib.error, urllib.parse, random, string, time, sys, os
HOST="https://pwabuilder-cloudapk.azurewebsites.net"
BASE="https://vanjamkrindo.github.io/excel-web"
OUT=os.path.join(os.path.dirname(os.path.abspath(__file__)),"apk-out")
os.makedirs(OUT,exist_ok=True)
def pw(n=12): return ''.join(random.choice(string.ascii_letters+string.digits+"!@#$%^&*") for _ in range(n))
opts={
 "packageId":"app.excelweb.spreadsheet","name":"Excel Web","launcherName":"Excel",
 "appVersion":"1.0.0","appVersionCode":1,
 "host":BASE+"/","startUrl":"/excel-web/index.html",
 "webManifestUrl":BASE+"/manifest.webmanifest","iconUrl":BASE+"/icons/icon-512.png",
 "themeColor":"#107c41","backgroundColor":"#ffffff","navigationColor":"#107c41",
 "display":"standalone","fallbackType":"customtabs","signingMode":"new",
 "splashScreenFadeOutDuration":300,"enableNotifications":False,"shortcuts":[],
 "signing":{"alias":"excelweb","fullName":"Alfan Andrianto","organization":"ExcelWeb",
   "organizationalUnit":"Dev","countryCode":"ID","keyStorePassword":pw(),"keyPassword":pw()}}
job=os.path.join(OUT,"job.json")
jid=tok=None
if os.path.exists(job):
    d=json.load(open(job)); jid,tok=d.get("id"),d.get("accessToken")
if not jid:
    req=urllib.request.Request(HOST+"/enqueuePackageJob",data=json.dumps(opts).encode(),
      headers={"Content-Type":"application/json"})
    try:
        with urllib.request.urlopen(req,timeout=90) as r: d=json.loads(r.read())
        jid,tok=d.get("id"),d.get("accessToken")
        json.dump({"id":jid,"token":tok},open(job,"w"))
        print("enqueued:",jid)
    except urllib.error.HTTPError as e:
        print("enqueue failed:",e.code,e.read()[:200].decode(errors="replace")); sys.exit(2)
for i in range(80):
    time.sleep(15)
    q=urllib.request.Request(HOST+"/getPackageJob?id="+urllib.parse.quote(jid),
      headers={"Authorization":"Bearer "+tok})
    try:
        with urllib.request.urlopen(q,timeout=60) as r: s=json.loads(r.read())
    except Exception as e:
        print("poll err",e); continue
    st=s.get("status"); print(f"[{i}] {st}",flush=True)
    if st=="Completed":
        dl=urllib.request.Request(HOST+"/downloadPackageZip?id="+urllib.parse.quote(jid),
          headers={"Authorization":"Bearer "+tok})
        with urllib.request.urlopen(dl,timeout=300) as r: data=r.read()
        p=os.path.join(OUT,"package.zip"); open(p,"wb").write(data)
        print("downloaded",len(data),"bytes ->",p); break
    if st in ("Failed","Error"):
        print("FAILED",json.dumps(s)[:600]); sys.exit(1)
else:
    print("still running"); sys.exit(3)
