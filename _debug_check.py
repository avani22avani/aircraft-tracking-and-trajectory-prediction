import traceback
from app import app

try:
    with app.test_client() as c:
        r = c.get('/')
        with open('debug_http.txt','w',encoding='utf-8') as f:
            f.write(f'STATUS={r.status_code}\n')
            f.write(r.get_data(as_text=True)[:4000])
except Exception:
    with open('debug_http.txt','w',encoding='utf-8') as f:
        f.write(traceback.format_exc())
