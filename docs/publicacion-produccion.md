# Publicación AIOS en Ubuntu / WSL

## Arquitectura y alcance

Windows compila y envía por SMB a \\172.19.130.163\Temp. El instalador se ejecuta
en Ubuntu de esa máquina y deja la aplicación en /opt/plantilla-aios-macro,
puerto 8084. No es la distribución WSL del portátil de desarrollo.

SMB no ejecuta comandos Linux. El envío elimina la copia manual por Escritorio
remoto; para activar la entrega ejecute el comando Bash mostrado por el script
en WSL de producción. Sin SSH, WinRM o una tarea autorizada en el servidor, este
último paso requiere una consola del servidor. No se declara desplegada una
entrega solo porque terminó su transferencia.

## 1. Preparar Ubuntu de producción (una vez)

Use la misma cuenta WSL que será propietaria del proceso y tenga acceso a sudo:

    sudo apt-get update
    sudo apt-get install openjdk-21-jre-headless tesseract-ocr tesseract-ocr-eng python3 rsync curl unzip util-linux fontconfig fonts-dejavu-core

Java 21 o posterior. No hace falta Maven, Excel ni ejecutar macros VBA en producción.
Se usan los generadores Java y las plantillas vacías incluidas en el JAR.
Se conserva la configuración Windows de desarrollo; el arranque Linux activa el
perfil prod, que exige una ruta de insumos explícita.

AIOS necesita acceso de red a Teradata 10.40.176.8, al servicio TRM y a las páginas
de cartas circulares de superfinanciera.gov.co para las comisiones con OCR.
El idioma eng de Tesseract debe estar instalado. La salud HTTP no certifica que
existan todos los insumos para todos los periodos.

## 2. Compilar y enviar desde PowerShell local

    cd D:\app\plantilla-aios-macro
    .\scripts\desplegar-produccion.ps1

Verifica que el código contenga origin/main, incrementa VERSION.RELEASE en el
POM y ejecuta las pruebas. El ZIP contiene solo el JAR, manifiesto y archivos de
operación: no incluye credenciales, insumos, documentos de diagnóstico ni output/.

Para reutilizar la entrega ya generada sin incrementar versión:

    .\scripts\desplegar-produccion.ps1 -Paquete 'D:\ruta\plantilla-aios-<version>-<fecha>.zip'

Cada entrega se envía a una subcarpeta aios-<fecha>-<id>. Se comprueba SHA-256;
si falla, queda .partial y no debe ejecutarse. El ZIP incluye su propio publicador
y validador, evitando mezclar versiones. -WhatIf no compila ni copia.

Temp es el nombre compartido, no una unidad. El valor WSL predeterminado supone
que apunta a C:\Temp. Si corresponde a D: u otra carpeta, especifique ambas rutas:

    .\scripts\desplegar-produccion.ps1 -Destino '\\172.19.130.163\Despliegues' -RutaWslDestino '/mnt/d/Despliegues'

El recurso debe existir y permitir escritura, lectura y renombrado. Si necesita
autenticación, en la misma consola use (la contraseña se solicita interactivamente):

    net use \\172.19.130.163\Temp /user:SUPERFIN\jcrojas * /persistent:no

## 3. Instalar en WSL de producción

Ejecute exactamente el comando mostrado al terminar el envío, por ejemplo:

    bash '/mnt/c/Temp/aios-<fecha>-<id>/publicar-produccion.sh' '/mnt/c/Temp/aios-<fecha>-<id>/plantilla-aios-<version>-<fecha>.zip'

Para validar el ZIP sin instalar ni pedir sudo:

    bash '/mnt/c/Temp/aios-<fecha>-<id>/publicar-produccion.sh' --verificar-paquete '/mnt/c/Temp/aios-<fecha>-<id>/plantilla-aios-<version>-<fecha>.zip'

En una primera instalación crea /opt/plantilla-aios-macro/.env con permisos 600
y se detiene antes de activar la aplicación. Edítelo directamente allí:

    nano /opt/plantilla-aios-macro/.env

Complete AIOS_DB_USER, AIOS_PASS y AIOS_INSUMOS_DIR (ruta Linux existente y legible,
por ejemplo /mnt/d/Datos/Pensiones/...; confirme la ruta real, no la del portátil).
Use comillas simples para valores con espacios o caracteres especiales. No copie
el .env a Git, al chat ni a la carpeta compartida. Es configuración Bash de confianza.
Puede indicar JAVA_CMD y TESSERACT_PATH absolutos si no usa los binarios del sistema.
Repita el comando de instalación.

El publicador conserva .env y las carpetas locales logs, target (caché y salidas),
insumos, plantillas y salidas_referencia. Los datos externos no se tocan. Prepara la
nueva entrega antes de detener la anterior; rechaza un puerto ocupado por otro proceso.
Si ya existe un AIOS arrancado manualmente o por otro gestor, deténgalo con su gestor
antes de migrar: no se matan procesos desconocidos ni se modifican otras aplicaciones.

Luego inicia el JAR sin root, comprueba /actuator/health (incluye BD) y verifica
que /actuator/info declare exactamente la versión instalada. Si falla, conserva
los archivos fallidos y restaura la instalación anterior en la medida en que pueda
detenerse el proceso nuevo. Los respaldos quedan en /opt/plantilla-aios-macro-backups.
Los intentos fallidos quedan en /opt/plantilla-aios-macro.failed.* o .new.*.

## 4. Operación y verificación

La interfaz AIOS incluye el botón **Volver al menú principal**, que abre Informes
Financieros conservando el servidor y protocolo usados en el navegador. Por defecto,
desde `http://localhost:8084/aios` vuelve a `http://localhost:8081/reportes`; desde
`http://172.19.130.163:8084/aios` vuelve a `http://172.19.130.163:8081/reportes`.
El enlace elimina los parámetros y el fragmento de la página AIOS.

Si cambia el puerto o la ruta del menú, configure `AIOS_MENU_PORT` y
`AIOS_MENU_PATH` en `/opt/plantilla-aios-macro/.env` y reinicie AIOS. Sus valores
predeterminados son `8081` y `/reportes`; la ruta debe comenzar por `/`.
El menú debe ser accesible en el mismo servidor y protocolo; estas variables no
modifican su despliegue ni las reglas de red.

    bash /opt/plantilla-aios-macro/scripts/manage-app.sh verificar
    bash /opt/plantilla-aios-macro/scripts/manage-app.sh status
    bash /opt/plantilla-aios-macro/scripts/manage-app.sh restart
    tail -f /opt/plantilla-aios-macro/logs/aios.log
    curl --fail http://127.0.0.1:8084/actuator/health
    curl --fail http://127.0.0.1:8084/actuator/info

Compruebe la interfaz /aios, el pie de versión y la generación de un periodo conocido
con insumos reales. En Ubuntu se distinguen mayúsculas y minúsculas en rutas.
La fecha del pie es la de compilación del JAR, en America/Bogota; copiar o reiniciar
no cambia esa fecha. Desde Windows servidor, pruebe http://localhost:8084/aios.

Acceso desde otros equipos a http://172.19.130.163:8084/aios depende del firewall y
la red WSL (NAT/reflejada o reenvío). No se abren puertos automáticamente.
El publicador no crea la tarea Windows. La configuración productiva usa
**Iniciar AIOS**, junto con **Iniciar Ubuntu** para mantener WSL activo.
El procedimiento y la verificación están al final de este documento.

## 5. Acceso desde otras máquinas: configuración de red

Esta sección documenta la configuración comprobada en producción el 1 de octubre
de 2026: Windows con WSL y reenvío de puertos hacia Ubuntu. Las direcciones internas
pueden cambiar; confirme los valores antes de ejecutar los comandos. No es una
receta para una instalación con red WSL reflejada.

### 5.1. Qué significa cada dirección

| Dirección | Pertenece a | Para qué se utiliza |
|---|---|---|
| `172.19.130.163` | Windows del servidor, en la red corporativa | Los usuarios acceden a `http://172.19.130.163:8084/aios`. |
| `172.28.36.220` | Ubuntu dentro de WSL | Windows reenvía las conexiones a AIOS en esta IP y el puerto 8084. |
| `172.28.32.1` | Adaptador virtual de Windows conectado a WSL | Es la IP de origen que Ubuntu recibe en las conexiones reenviadas por Windows; se autoriza en UFW. |
| `127.0.0.1` | Interfaz local del entorno donde se usa | Permite probar AIOS dentro de Ubuntu. El acceso por localhost en Windows no demuestra acceso desde la red. |
| `0.0.0.0` | Dirección de escucha, no una máquina | Indica que se aceptan conexiones por todas las interfaces IPv4 del entorno. No se escribe como destino en el navegador. |

Windows tiene una dirección en la red corporativa y otra en la red virtual de WSL.
El puerto 8084 identifica el servicio AIOS dentro de cada dirección.

```text
Equipo del usuario
  -> Windows del servidor: 172.19.130.163:8084
  -> portproxy de Windows (origen hacia WSL: 172.28.32.1)
  -> Ubuntu WSL: 172.28.36.220:8084
  -> Aplicación AIOS
```

### 5.2. Permitir que AIOS escuche en la interfaz de WSL

En **Ubuntu de producción**, edite el archivo:

```bash
nano /opt/plantilla-aios-macro/.env
```

Compruebe estos valores:

```bash
AIOS_PORT=8084
AIOS_BIND_ADDRESS=0.0.0.0
```

El perfil `prod` usa estas variables para configurar el puerto y la dirección de
escucha. `0.0.0.0` permite recibir conexiones dirigidas a la IP de Ubuntu; los
firewalls siguen controlando quién puede conectarse. Si cambió estos valores con
la aplicación en ejecución, reinicie únicamente AIOS:

```bash
bash /opt/plantilla-aios-macro/scripts/manage-app.sh restart
```

Verifique el proceso, la escucha y la página:

```bash
bash /opt/plantilla-aios-macro/scripts/manage-app.sh status
ss -lntp 'sport = :8084'
curl --noproxy '*' -sS -o /dev/null --max-time 10 \
  -w 'HTTP local: %{http_code}\n' http://127.0.0.1:8084/aios
hostname -I
```

Se espera `AIOS activo: <PID>`, una escucha en `*:8084` o `0.0.0.0:8084`, y HTTP
`200`. `hostname -I` puede mostrar varias IP; en esta instalación la interfaz WSL
usa `172.28.36.220`. No seleccione una IP de otra red, como un puente de contenedores.

### 5.3. Reenviar el puerto de Windows hacia Ubuntu

En **PowerShell como administrador del servidor 172.19.130.163**, consulte primero
la configuración existente:

```powershell
netsh interface portproxy show all
```

Si aún no existe el reenvío de AIOS, agréguelo con la IP actual de WSL:

```powershell
netsh interface portproxy add v4tov4 listenaddress=0.0.0.0 listenport=8084 connectaddress=172.28.36.220 connectport=8084
```

Este comando hace que Windows escuche en el 8084 y reenvíe las conexiones al 8084
de Ubuntu. No modifica los reenvíos existentes de 8081, 8083 o 8087.

Si la entrada ya existe y solo cambió la IP de WSL, actualícela con el valor real:

```powershell
netsh interface portproxy set v4tov4 listenaddress=0.0.0.0 listenport=8084 connectaddress=172.28.36.220 connectport=8084
```

### 5.4. Permitir la entrada en el firewall de Windows

En la misma **PowerShell como administrador del servidor**, consulte la regla:

```powershell
Get-NetFirewallRule -DisplayName 'Allow AIOS WSL 8084' -ErrorAction SilentlyContinue |
    Select-Object DisplayName,Enabled,Direction,Action,Profile
```

Si no existe, cree la regla utilizada en esta instalación:

```powershell
New-NetFirewallRule -DisplayName 'Allow AIOS WSL 8084' -Direction Inbound -Action Allow -Protocol TCP -LocalPort 8084 -Profile Any
```

Esta regla permite conexiones TCP entrantes al 8084 de Windows en todos los
perfiles de red. Si TI requiere limitar los equipos de origen o perfiles, adapte
el alcance a la red autorizada. No es necesario crear reglas duplicadas.

### 5.5. Permitir la entrada en el firewall de Ubuntu (UFW)

En **Ubuntu**, consulte las reglas:

```bash
sudo ufw status verbose
```

En producción UFW estaba activo, con entrada denegada por defecto y permisos para
8081 y 8087, pero no para 8084. Esto permitía abrir AIOS dentro de Ubuntu y bloqueaba
la conexión desde Windows. Se resolvió con:

```bash
sudo ufw allow from 172.28.32.1 to any port 8084 proto tcp
```

La regla permite únicamente el origen Windows de la red virtual de WSL. La IP
`172.28.32.1` se comprobó en `SourceAddress` al ejecutar desde Windows
`Test-NetConnection 172.28.36.220 -Port 8084`. Use el origen real de su instalación.
La regla se aplica inmediatamente; no requiere reiniciar AIOS ni desactivar UFW.
Si UFW está inactivo, no lo active solo para añadir esta regla; revise el filtrado
que realmente utilice el servidor.

### 5.6. Verificar el recorrido completo

En **PowerShell del servidor**, compruebe la conexión directa hacia Ubuntu:

```powershell
Test-NetConnection 172.28.36.220 -Port 8084
curl.exe --noproxy "*" --fail --max-time 10 http://172.28.36.220:8084/aios -o NUL
Get-NetTCPConnection -State Listen -LocalPort 8084 |
    Select-Object LocalAddress,LocalPort,OwningProcess
```

Se espera `TcpTestSucceeded : True`, una respuesta HTTP satisfactoria y una escucha
de Windows en `0.0.0.0:8084`. Una escucha solo en `::1` corresponde a acceso local.

En **PowerShell de otro equipo de la red**, compruebe el acceso al servidor:

```powershell
Test-NetConnection 172.19.130.163 -Port 8084
curl.exe --noproxy "*" --fail --max-time 10 http://172.19.130.163:8084/aios -o NUL
```

Abra después `http://172.19.130.163:8084/aios` en el navegador y genere un periodo
conocido. Un resultado TCP positivo solo confirma que el puerto acepta conexiones:
`portproxy` puede aceptarlas aunque no consiga comunicarse con Ubuntu. La prueba
HTTP verifica un paso adicional; la generación comprueba las fuentes de datos.

| Resultado | Qué revisar |
|---|---|
| AIOS no responde por localhost dentro de Ubuntu | Proceso AIOS y `logs/aios.log`. |
| Responde por localhost, pero no por la IP de WSL dentro de Ubuntu | Dirección de escucha, IP actual y filtrado local. |
| Responde por ambas direcciones dentro de Ubuntu, pero Windows no alcanza el 8084 de WSL | UFW y, si no explica el bloqueo, firewall de Hyper-V. |
| Windows alcanza Ubuntu, pero otro equipo no alcanza el puerto del servidor | Escucha de portproxy, firewall de Windows y red corporativa. |
| TCP desde otro equipo funciona, pero HTTP falla | Destino de portproxy y respuesta HTTP directa desde Windows hacia WSL. |

### 5.7. Después de un reinicio

Las reglas quedan guardadas, pero las IP internas pueden cambiar. Confirme la IP
de Ubuntu con `hostname -I` y el origen Windows con `Test-NetConnection`; actualice
el destino de `portproxy` y la regla de UFW si corresponde. Si cambia el origen,
añada primero la regla nueva y elimine la anterior solo después de verificarla
con `sudo ufw status numbered`.

Esta configuración de red no programa el arranque automático de WSL ni de AIOS.
Configure por separado la tarea indicada en la sección 4 si debe recuperarse el
servicio al reiniciar Windows. No reinicie WSL completo para un cambio de AIOS,
pues puede interrumpir las otras aplicaciones alojadas allí.

Referencia: [Red y acceso a aplicaciones de WSL (Microsoft)](https://learn.microsoft.com/en-us/windows/wsl/networking).


## Arranque automático en Windows de producción (2 de octubre de 2026)

Producción: **DPFUENTESW**, IP corporativa `172.19.130.163`. Los proyectos locales
`D:\app\ConsultaTRMWeb` y `D:\app\plantilla-aios-macro` son fuentes de desarrollo;
las aplicaciones instaladas se ejecutan en Ubuntu WSL bajo `/opt`.

La tarea Windows **Iniciar Ubuntu** inicia InformesFinancieros y mantiene Ubuntu
activo para las tres aplicaciones. Se ejecuta como **SUPERFIN\jcrojas**, aunque
no haya sesión abierta (contraseña guardada), al iniciar Windows con un minuto
de retraso. Su programa es `C:\Windows\System32\wsl.exe` y sus argumentos son:

```text
-d Ubuntu -u jcrojas --exec /bin/bash -lc "/opt/informes-financieros/scripts/manage-app.sh start && exec /usr/bin/sleep infinity"
```

En **Propiedades > Configuración**, desmarcar el límite de duración y seleccionar
**No iniciar una instancia nueva**. El panel inferior del Programador es de solo
lectura. Esta tarea permanece **En ejecución**. No terminarla ni modificarla para
iniciar AIOS o TRM: sostiene WSL compartido. `sleep infinity` no supervisa los
procesos de las aplicaciones.

Administrar tareas desde PowerShell con permisos administrativos en producción:

```powershell
hostname
whoami
Get-ScheduledTask -TaskName "Iniciar Ubuntu" | Select-Object TaskPath,TaskName,State
(Get-ScheduledTask -TaskName "Iniciar Ubuntu").Principal | Format-List UserId,LogonType
(Get-ScheduledTask -TaskName "Iniciar Ubuntu").Actions | Format-List Execute,Arguments
```

Esperado: equipo `DPFUENTESW`, tarea `Running` y cuenta propietaria de WSL.
Desde PowerShell de **SUPERFIN\jcrojas**, `wsl --list --running` debe listar
Ubuntu. La cuenta `DPFUENTESW\Administrador` puede tener otra distribución Ubuntu
detenida; `-u jcrojas` solo cambia el usuario Linux, no la cuenta Windows.

Después de configurar el inicio, cerrar las terminales Ubuntu, esperar dos minutos
y verificar HTTP desde Windows. Repetir tras el próximo reinicio programado,
esperando dos o tres minutos, **sin abrir Ubuntu manualmente**. Así se comprueba
el arranque automático, no un arranque provocado por la prueba.

```powershell
curl.exe --noproxy "*" -I --max-time 10 http://127.0.0.1:8081/reportes
curl.exe --noproxy "*" -I --max-time 10 http://127.0.0.1:8084/aios
curl.exe --noproxy "*" -I --max-time 10 http://127.0.0.1:8087/
```

Esperado: HTTP 200 para cada aplicación instalada. Si Ubuntu responde pero Windows
no, comparar la IP actual (`hostname -I` dentro de Ubuntu) con
`netsh interface portproxy show all`. No cambiar IP ni firewall sin localizar el
fallo. Si no hay distribuciones activas bajo la cuenta correcta, revisar primero
la tarea persistente. No reiniciar WSL como diagnóstico inicial: afecta las tres
aplicaciones. Estos pasos no publican ni actualizan los binarios instalados.

### Configurar la tarea Iniciar AIOS

Primero abrir Ubuntu como `jcrojas` y verificar la instalación existente:

```bash
bash /opt/plantilla-aios-macro/scripts/manage-app.sh verificar
bash /opt/plantilla-aios-macro/scripts/manage-app.sh start
```

En el Programador de tareas de producción, crear **Iniciar AIOS** si no existe:

- Cuenta: `SUPERFIN\jcrojas`; ejecutar aunque no haya sesión iniciada, con contraseña.
- Desencadenador: al iniciar el sistema, retraso de dos minutos.
- Programa: `C:\Windows\System32\wsl.exe`.
- Argumentos: la línea siguiente.
- Permitir inicio a petición; no iniciar una instancia nueva; límite de duración desmarcado.

```text
-d Ubuntu -u jcrojas --exec /bin/bash /opt/plantilla-aios-macro/scripts/manage-app.sh start
```

Guardar; ejecutar con clic derecho > Ejecutar o PowerShell administrativa:

```powershell
Start-ScheduledTask -TaskName "Iniciar AIOS"
Get-ScheduledTaskInfo -TaskName "Iniciar AIOS" | Format-List LastRunTime,LastTaskResult
```

Esperar 30 a 60 segundos antes de la prueba HTTP. Esta tarea puede terminar en
**Listo**, con resultado 0: el script deja Java en segundo plano mediante `nohup`
y `&`. Eso no mantiene WSL activo por sí solo; depende de **Iniciar Ubuntu**.
Un resultado 0 no sustituye la verificación HTTP ni supervisa fallos posteriores.

Diagnóstico dentro de Ubuntu:

```bash
bash /opt/plantilla-aios-macro/scripts/manage-app.sh status
tail -n 80 /opt/plantilla-aios-macro/logs/aios.log
curl --fail --max-time 10 http://127.0.0.1:8084/actuator/health
```

No imprimir ni compartir `.env`: contiene configuración sensible. Para TRM no
se crea otra tarea Windows: se habilita `consulta-trm.service` dentro de Ubuntu.
