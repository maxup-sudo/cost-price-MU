"""Production entrypoint. HTTPS terminates at your trusted reverse proxy."""
import os
if len(os.environ.get('MAXUP_SECRET_KEY',''))<32:
    raise RuntimeError('Configure MAXUP_SECRET_KEY with at least 32 random characters')
if os.environ.get('MAXUP_HTTPS')!='1':
    raise RuntimeError('Production requires HTTPS and MAXUP_HTTPS=1')
from app import app
