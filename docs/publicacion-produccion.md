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
El script no programa arranque después de reiniciar Windows. Si se requiere,
configure con TI una tarea de Windows bajo la cuenta propietaria de la distribución:

    wsl.exe -d <Distribucion> -u <UsuarioWSL> -- bash /opt/plantilla-aios-macro/scripts/manage-app.sh start

No cambie la tarea existente de Informes Financieros: AIOS usa otra carpeta y puerto.
