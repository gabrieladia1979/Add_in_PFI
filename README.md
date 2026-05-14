# Add-in PFI

Skeleton inicial del add-in de Outlook para el proyecto de ciberdefensa con XAI.

## Requisitos
- Node.js 18+
- Outlook Desktop o Outlook Web (con sideload habilitado)

## Instalacion
1. npm install

## Ejecutar
- npm run start

## Notas
- El panel se carga en lectura de mensajes (Message Read).
- El manifest principal esta en manifest.xml.
 - Sideload manual: en Outlook, agrega manifest.xml desde Complementos.
 - Si tu entorno no confia en el certificado local, acepta el certificado del dev server.
