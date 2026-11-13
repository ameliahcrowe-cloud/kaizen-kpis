import re,json
EX='https://script.google.com/macros/s/AKfycbygNx44F_JhoaR8DXoKocBxCHLQavLw0o85jA2_JmkhFzwx6cdx1nOG9sYduOMLVkdxOw/exec'
def page(title,body):
    return ('<!doctype html><html lang="en"><head><meta charset="utf-8">'
     '<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">'
     f'<title>{title}</title><link rel="icon" href="icon-192.png"><link rel="apple-touch-icon" href="icon-512.png">'
     '<link rel="manifest" href="manifest.webmanifest"><meta name="theme-color" content="#151515">'
     '<meta name="apple-mobile-web-app-capable" content="yes"><meta name="mobile-web-app-capable" content="yes">'
     f'<meta name="apple-mobile-web-app-title" content="{title}">'
     '</head><body>'+body+'</body></html>')
I=open('midas-src/Index.html').read()
I=re.sub(r"const HOURLY_URL='[^']*';","const HOURLY_URL='hourly.html';",I)
open('m/index.html','w').write(page('Midas KPIs',I))
open('m/hourly.html','w').write(page('Midas Hourly',open('midas-src/Hourly.html').read()))
open('m/manifest.webmanifest','w').write(json.dumps({"name":"Midas KPIs","short_name":"Midas","start_url":"./","display":"standalone","background_color":"#ffffff","theme_color":"#151515","icons":[{"src":"icon-192.png","sizes":"192x192","type":"image/png"},{"src":"icon-512.png","sizes":"512x512","type":"image/png"}]}))
print('built')
