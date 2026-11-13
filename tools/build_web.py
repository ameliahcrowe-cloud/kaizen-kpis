import re,json,shutil,os
def page(title,body,theme):
    return ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
     '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">'
     f'<title>{title}</title><link rel="icon" href="icon-192.png"><link rel="apple-touch-icon" href="icon-512.png">'
     f'<link rel="manifest" href="manifest.webmanifest"><meta name="theme-color" content="{theme}">'
     '<meta name="apple-mobile-web-app-capable" content="yes"><meta name="mobile-web-app-capable" content="yes">'
     f'<meta name="apple-mobile-web-app-title" content="{title}">'
     '</head><body>'+body+'</body></html>')
def build(src,out,name,short,theme,icons):
    os.makedirs(out,exist_ok=True)
    for a,b in icons: shutil.copy(a,os.path.join(out,b))
    I=open(src+'/Index.html').read()
    I=re.sub(r"const HOURLY_URL='[^']*';","const HOURLY_URL='hourly.html';",I)
    open(out+'/index.html','w').write(page(name+' KPIs',I,theme))
    open(out+'/hourly.html','w').write(page(name+' Hourly',open(src+'/Hourly.html').read(),theme))
    open(out+'/manifest.webmanifest','w').write(json.dumps({"name":name+" KPIs","short_name":short,"start_url":"./","display":"standalone","background_color":"#ffffff","theme_color":theme,"icons":[{"src":"icon-192.png","sizes":"192x192","type":"image/png"},{"src":"icon-512.png","sizes":"512x512","type":"image/png"}]}))
build('midas-src','m','Midas','Midas','#151515',[('icons/midas-192.png','icon-192.png'),('icons/midas-512.png','icon-512.png')])
build('orexis-src','o','Orexis','Orexis','#424965',[('orexis/icon-192.png','icon-192.png'),('orexis/icon-512.png','icon-512.png')])
print('built')
