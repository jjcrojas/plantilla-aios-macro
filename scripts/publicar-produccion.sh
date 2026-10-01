#!/usr/bin/env bash
set -Eeuo pipefail
APP=/opt/plantilla-aios-macro
BACKUPS=/opt/plantilla-aios-macro-backups
MODE=install
[[ "${1:-}" != --verificar-paquete ]] || { MODE=verify; shift; }
ZIP="${1:-}"
[[ -f "$ZIP" ]] || { echo 'Uso: bash publicar-produccion.sh [--verificar-paquete] archivo.zip' >&2; exit 2; }
HERE="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
for cmd in python3 mktemp; do command -v "$cmd" >/dev/null || exit 1; done
TEMP="$(mktemp -d /tmp/aios-publicacion.XXXXXXXX)"
trap '[[ "$TEMP" == /tmp/aios-publicacion.* ]] && rm -rf -- "$TEMP"' EXIT
VERSION="$(python3 "$HERE/validar-paquete.py" "$ZIP" "$TEMP")"
echo "Paquete AIOS $VERSION valido."
[[ "$MODE" != verify ]] || exit 0
for cmd in sudo rsync curl flock; do command -v "$cmd" >/dev/null || { echo "Falta $cmd" >&2; exit 1; }; done
[[ $(id -u) != 0 ]] || { echo 'Ejecute como usuario WSL de la aplicacion, no como root.' >&2; exit 1; }
sudo -v
sudo install -d -m 755 "$APP" "$BACKUPS"
[[ ! -L "$APP" && ! -L "$BACKUPS" ]] || { echo 'No se permiten enlaces en las rutas de instalacion.' >&2; exit 1; }
sudo touch /opt/plantilla-aios-macro.deploy.lock
sudo chown "$(id -u):$(id -g)" /opt/plantilla-aios-macro.deploy.lock
exec 8>/opt/plantilla-aios-macro.deploy.lock
flock -n 8 || { echo 'Otra publicacion AIOS esta en curso.' >&2; exit 1; }
if [[ ! -f "$APP/.env" ]]; then
  sudo install -m 600 -o "$(id -u)" -g "$(id -g)" "$TEMP/.env.example" "$APP/.env"
  echo "Configure $APP/.env con rutas y credenciales reales y repita el comando. No se instalo ni detuvo la aplicacion."
  exit 2
fi
[[ -r "$APP/.env" ]] || { echo 'Use el usuario WSL propietario de .env.' >&2; exit 1; }
set -a
source "$APP/.env"
set +a
PORT="${AIOS_PORT:-8084}"
[[ "$PORT" =~ ^[0-9]+$ ]] || exit 1
STAMP="$(date +%Y%m%d-%H%M%S)-$$"
NEW="/opt/plantilla-aios-macro.new.$STAMP"
BACKUP="$BACKUPS/$STAMP"
FAILED="/opt/plantilla-aios-macro.failed.$STAMP"
OLD_RUNNING=0
STOPPED=0
MOVED=0
ACTIVATED=0
rollback() {
  trap - ERR INT TERM
  set +e
  echo 'Publicacion fallida; restaurando el estado anterior.' >&2
  if [[ "$ACTIVATED" == 1 ]]; then
    bash "$APP/scripts/manage-app.sh" stop
    if bash "$APP/scripts/manage-app.sh" status >/dev/null 2>&1; then
      echo "No se pudo detener AIOS; respaldo intacto en $BACKUP. Requiere revision manual." >&2
      exit 1
    fi
    sudo mv "$APP" "$FAILED"
  fi
  if [[ "$MOVED" == 1 ]]; then sudo mv "$BACKUP" "$APP"; fi
  if [[ "$STOPPED" == 1 && "$OLD_RUNNING" == 1 ]]; then bash "$APP/scripts/manage-app.sh" start; fi
  echo "Revise logs; paquete fallido en $FAILED o $NEW, respaldo en $BACKUP." >&2
  exit 1
}
trap rollback ERR INT TERM
sudo install -d -o "$(id -u)" -g "$(id -g)" "$NEW"
rsync -a "$TEMP/" "$NEW/"
cp -p "$APP/.env" "$NEW/.env"
chmod 600 "$NEW/.env"
for dir in logs target insumos plantillas salidas_referencia; do
  if [[ -d "$APP/$dir" ]]; then rsync -a "$APP/$dir/" "$NEW/$dir/"; fi
done
chmod +x "$NEW/scripts/"*.sh
bash "$NEW/scripts/manage-app.sh" verificar
if [[ -f "$APP/scripts/manage-app.sh" ]]; then
  if bash "$APP/scripts/manage-app.sh" status >/dev/null 2>&1; then OLD_RUNNING=1; fi
  bash "$APP/scripts/manage-app.sh" stop
  STOPPED=1
fi
python3 - "$PORT" <<'PY'
import socket, sys
with socket.socket() as s:
    s.bind(('0.0.0.0', int(sys.argv[1])))
PY
sudo mv "$APP" "$BACKUP"
MOVED=1
sudo mv "$NEW" "$APP"
ACTIVATED=1
bash "$APP/scripts/manage-app.sh" start
healthy=0
for _ in {1..90}; do
  if bash "$APP/scripts/manage-app.sh" status >/dev/null 2>&1 &&
     curl -fsS --max-time 3 "http://127.0.0.1:$PORT/actuator/health" 2>/dev/null | python3 -c 'import json,sys; sys.exit(json.load(sys.stdin).get("status") != "UP")' 2>/dev/null &&
     curl -fsS --max-time 3 "http://127.0.0.1:$PORT/actuator/info" 2>/dev/null | python3 -c 'import json,sys; sys.exit(json.load(sys.stdin).get("build",{}).get("version") != sys.argv[1])' "$VERSION" 2>/dev/null; then
    healthy=1; break
  fi
  sleep 2
done
[[ "$healthy" == 1 ]] || { echo 'Salud o version incorrectas.' >&2; false; }
trap - ERR INT TERM
echo "AIOS $VERSION instalado y verificado en $APP, puerto $PORT. Respaldo: $BACKUP"
echo 'Verifique ademas la generacion de un periodo con insumos reales.'
