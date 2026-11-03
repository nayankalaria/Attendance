ATTENDANCE APP - RENDER REMOTE WHATSAPP BROWSER (noVNC)

Files in this package:
- app.py: your attendance application, changed to launch visible Chromium and
  includes a link to open the remote browser in a new tab.
- Dockerfile: installs Xvfb, x11vnc, noVNC, websockify, nginx and supervisor.
- nginx.conf: routes the public Render port to Flask and noVNC.
- supervisord.conf: runs the virtual desktop, VNC, noVNC, Flask and nginx.

DEPLOY STEPS
1. Replace app.py in your GitHub repository with this package's app.py.
2. Add Dockerfile, nginx.conf and supervisord.conf to the repository root.
3. Keep your existing requirements.txt (including Playwright 1.64.1).
4. In Render, keep the service as a Docker Web Service and set Dockerfile path to ./Dockerfile.
5. Add/confirm these environment variables:
   APP_USERNAME = your attendance login username
   APP_PASSWORD = a strong attendance login password
   SECRET_KEY = a long random secret
   VNC_PASSWORD = a separate strong VNC password, at least 8 characters
   DATA_DIR = /data
   COOKIE_SECURE = true
6. If your Render plan supports a persistent disk, mount it at /data so the
   WhatsApp profile and attendance data survive restarts/redeploys.
7. Deploy. Use Clear build cache & deploy if changing the Docker setup.

USAGE
1. Sign in to the attendance site.
2. Click Connect WhatsApp once.
3. Click "Open remote WhatsApp browser in a new tab".
4. Enter the VNC password when prompted.
5. The remote Chromium desktop should appear. Scan WhatsApp's QR using your
   phone: WhatsApp > Linked devices > Link a device.
6. Keep the Render service running while using WhatsApp automation.

SECURITY
- Never share VNC_PASSWORD, APP_PASSWORD, SECRET_KEY, QR codes, or session data.
- VNC is exposed through the Render HTTPS endpoint; use a strong separate
  VNC_PASSWORD. VNC password authentication is legacy and may effectively use
  only the first 8 characters, so use 8 distinct random characters at minimum.
- Do not expose port 5900 directly; this setup binds x11vnc to localhost and
  exposes the browser only through noVNC/websockify on the Render web port.
- A persistent disk may not be available on every Render plan. Without one,
  WhatsApp login and attendance data may be lost after restarts/redeploys.
