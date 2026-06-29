# pbMailKit — Enviar correo por SMTP desde PowerBuilder ✉️

![PowerBuilder](https://img.shields.io/badge/PowerBuilder-2025-2D6CDF?style=flat-square)
![.NET](https://img.shields.io/badge/.NET-10-512BD4?style=flat-square&logo=dotnet&logoColor=white)
![SMTP](https://img.shields.io/badge/SMTP-TLS%20%7C%20Adjuntos-00A98F?style=flat-square)
![Blog](https://img.shields.io/badge/blog-rsrsystem-FF5722?style=flat-square&logo=blogger&logoColor=white)

## 📋 ¿Qué es esto?

Un ejemplo PowerBuilder para **enviar correos electrónicos por SMTP**: con autenticación,
cifrado **TLS/SSL**, varios destinatarios (To/CC/BCC), cuerpo en texto o **HTML**, prioridad,
acuse de recibo y **adjuntos**. Vamos, lo que hace falta para mandar correo "de verdad" desde una
aplicación de gestión.

¿Y la novedad? Que en **PowerBuilder 2025 GA (Build 3683)** ya **no hace falta una librería
externa**: empleamos la nueva clase **nativa `SMTPClient`** de PowerBuilder. En el ejemplo la
extendemos en un objeto **`nvo_smtpclient`** (heredado de `smtpclient`) con un puñado de constantes
cómodas para el protocolo seguro (TLS 1.0–1.3), la prioridad, los juegos de caracteres y los
distintos *reset*. Así montáis el correo y lo enviáis con muy poco código y sin DLLs de terceros
rondando por la carpeta.

## 🔗 Motor .NET

En la versión actual el motor es la **clase nativa `SMTPClient` de PowerBuilder 2025**, así que no
hay que desplegar ninguna DLL ni cargar nada como `dotnetobject`.

Si necesitas (o quieres estudiar) la versión que se apoyaba en una **librería .NET propia**, está
la librería **`MailKitNetSmtp`** (clase `MailKitSmptRSR`), que envuelve **MailKit + MimeKit** para
hacer exactamente lo mismo —conexión `Auto`/`SslOnConnect`/`StartTls`, autenticación, adjuntos,
HTML, prioridad y acuse de recibo— y se consumía desde PowerBuilder como `dotnetobject`:

- **Código fuente** en `Blog\Net10\MailKitNetSmtp` (antes estaba en `Net8`); se recompila/despliega
  con el script **`desplegar_dotnet.bat`**.
- Repo del proyecto .NET (Visual Studio 2022): <https://github.com/rasanfe/MailKitNetSmtp>

## 🛠️ Requisitos

- **PowerBuilder 2025 GA (Build 3683)** para abrir y compilar la solución (trae la clase nativa
  `SMTPClient`).
- Una cuenta SMTP a la que conectaros (servidor, puerto, usuario y contraseña). Para Gmail y
  similares, recordad usar **contraseña de aplicación** si tenéis 2FA.

## ▶️ Cómo probarlo

1. Clona el repo y abre `pbmailkit.pbsln` con PowerBuilder 2025.
2. Compila (Full Build) y ejecuta.
3. Rellena los datos de tu servidor SMTP (servidor, puerto, usuario, contraseña, tipo de conexión).
4. Escribe remitente, destinatario, asunto y cuerpo, añade algún adjunto y pulsa enviar.

## 🗂️ Versiones anteriores (repo archivo)

En el repositorio **archivo** tienes versiones previas de este ejemplo para distintas builds de
PowerBuilder (las que usaban la librería MailKit por `dotnetobject`):

- <https://github.com/rasanfe/archivo/tree/main/PowerBuilder_115_b2506/pbMailkit115>
- <https://github.com/rasanfe/archivo/tree/main/PowerBuilder_126_b3506/pbMailkit126>
- <https://github.com/rasanfe/archivo/tree/main/PowerBuilder_2019_b2779/pbMailkit2019>
- <https://github.com/rasanfe/archivo/tree/main/PowerBuilder_2021_b1509/pbMailkit2021>
- <https://github.com/rasanfe/archivo/tree/main/PowerBuilder_2022_b1878/pbMailKit2022>
- <https://github.com/rasanfe/archivo/tree/main/PowerBuilder_2022_b3359/pbMailKit2022>
- <https://github.com/rasanfe/archivo/tree/main/PowerBuilder_2022_b3397/pbMailKit>

## 🔗 Repo PowerBuilder

<https://github.com/rasanfe/pbMailKit>

---

> ¡Nos vemos en el próximo artículo! Y recuerda: en PowerBuilder, los límites solo están en nuestra imaginación. 🚀

📨 **Blog:** <https://rsrsystem.blogspot.com/>
