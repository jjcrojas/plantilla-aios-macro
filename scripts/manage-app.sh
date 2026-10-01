#!/usr/bin/env bash
set -euo pipefail
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
[[ -f "$ROOT/.env" ]] || { echo "Falta $ROOT/.env" >&2; exit 1; }
set -a
source "$ROOT/.env"
set +a
JAVA_CMD="${JAVA_CMD:-java}"
PID_FILE="$ROOT/logs/aios.pid"
mkdir -p "$ROOT/logs"
exec 9>"$ROOT/logs/manage.lock"
flock -n 9 || { echo 'Hay otra operacion de arranque/parada en curso.' >&2; exit 1; }
fail() { echo "[AIOS] $*" >&2; exit 1; }
jar_name() { sed -n 's/^jar.name=//p' "$ROOT/publicacion-manifest.properties" | head -n 1; }
running() {
  [[ -f "$PID_FILE" ]] || return 1
  pid="$(cat "$PID_FILE")"
  [[ "$pid" =~ ^[0-9]+$ && -r "/proc/$pid/cmdline" ]] || return 1
  kill -0 "$pid" 2>/dev/null && tr '\0' '\n' < "/proc/$pid/cmdline" | grep -Fxq "$ROOT/$(jar_name)"
}
preflight() {
  [[ -n "${AIOS_DB_USER:-}" && -n "${AIOS_PASS:-}" ]] || fail 'Configure AIOS_DB_USER y AIOS_PASS en .env.'
  [[ -n "${AIOS_INSUMOS_DIR:-}" && -d "$AIOS_INSUMOS_DIR" && -r "$AIOS_INSUMOS_DIR" ]] || fail 'AIOS_INSUMOS_DIR debe ser un directorio legible en Ubuntu.'
  [[ "${AIOS_PORT:-8084}" =~ ^[0-9]+$ ]] || fail 'Puerto invalido.'
  (( ${AIOS_PORT:-8084} > 1023 && ${AIOS_PORT:-8084} < 65536 )) || fail 'Puerto fuera de rango.'
  command -v "$JAVA_CMD" >/dev/null || fail 'Falta Java 21 o posterior.'
  local major
  major="$("$JAVA_CMD" -version 2>&1 | sed -nE 's/.*version "([0-9]+).*/\1/p' | head -n1)"
  [[ "$major" =~ ^[0-9]+$ ]] && (( major >= 21 )) || fail 'AIOS requiere Java 21 o posterior.'
  [[ -x "${TESSERACT_PATH:-/usr/bin/tesseract}" ]] || fail 'Falta Tesseract OCR. Configure TESSERACT_PATH.'
  "${TESSERACT_PATH:-/usr/bin/tesseract}" --list-langs 2>/dev/null | grep -Fxq eng || fail 'Falta el idioma eng de Tesseract.'
  [[ -f "$ROOT/$(jar_name)" ]] || fail 'Falta el JAR del manifiesto.'
}
stop_app() {
  if running; then
    kill -TERM "$pid"
    for _ in {1..60}; do running || break; sleep 1; done
    running && fail 'El proceso no termino. No se forzo su cierre ni se reemplazaron archivos.'
  elif [[ -f "$PID_FILE" ]] && kill -0 "$(cat "$PID_FILE")" 2>/dev/null; then
    fail 'El PID corresponde a otro proceso; no se detuvo.'
  fi
  rm -f -- "$PID_FILE"
}
start_app() {
  running && { echo "AIOS ya activo: $pid"; return; }
  preflight
  local -a opts=()
  read -r -a opts <<< "${JAVA_OPTS:--Xms256m -Xmx2g -Djava.awt.headless=true}"
  cd "$ROOT"
  nohup "$JAVA_CMD" "${opts[@]}" -jar "$ROOT/$(jar_name)" --spring.profiles.active=prod >> "$ROOT/logs/aios.log" 2>&1 8>&- 9>&- < /dev/null &
  echo "$!" > "$PID_FILE"
  sleep 2
  running || fail "Fallo el arranque; consulte $ROOT/logs/aios.log"
}
case "${1:-}" in
  verificar) preflight; echo 'Requisitos locales correctos (conexion a Teradata se verifica al arrancar).' ;;
  start) start_app ;;
  stop) stop_app ;;
  restart) stop_app; start_app ;;
  status) running && echo "AIOS activo: $pid" ;;
  *) fail 'Uso: manage-app.sh {verificar|start|stop|restart|status}' ;;
esac
