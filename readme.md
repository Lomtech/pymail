# 📧 Pymailer – Serien-E-Mails aus Excel versenden

Dieses Tool ermöglicht es, Serien-E-Mails direkt aus einer Excel-Datei zu versenden.  
Die Empfängerliste, der Betreff und der E-Mail-Text stammen aus der Excel-Datei.  
Das E-Mail-Layout wird mit einer HTML-Vorlage gestaltet.  

---

## 🔹 Voraussetzungen

- **Windows-PC** (benötigt Outlook)  
- **Excel-Datei** im gleichen Ordner wie das Programm (`Kundenliste.xlsx`)  

---

## 🔹 Installation (mit Python)

1. Projekt herunterladen:  
   ```bash
   git clone https://github.com/deinname/pymailer.git
   cd pymailer
Abhängigkeiten installieren:

bash
Kopieren
pip install -r requirements.txt
Excel-Datei Kundenliste.xlsx im Projektordner anpassen
(Spalten: E-Mail, Betreff, Name, …).

HTML-Vorlage mail_template.html bei Bedarf anpassen.

Starten:

bash
Kopieren
python mailer.py
🔹 Installation (ohne Python)
Die Datei mailer.exe aus dem Ordner dist/ herunterladen.
(Alternativ in den GitHub Releases verfügbar).

In einen beliebigen Ordner legen – zusammen mit:

Kundenliste.xlsx

mail_template.html

Per Doppelklick starten.

🔹 Excel-Datei (Beispiel Kundenliste.xlsx)
Name	Email	Betreff	Platzhalter1	Platzhalter2
Max Mustermann	max@example.com	Willkommen bei uns!	Lieber Max	Produkt X
Erika Muster	erika@example.com	Ihr Angebot	Liebe Erika	Produkt Y

➡️ Jeder Platzhalter ({{ Platzhalter1 }} etc.) kann in der HTML-Vorlage verwendet werden.

🔹 HTML-Vorlage (Beispiel mail_template.html)
html
Kopieren
<html>
  <body>
    <p>{{ Platzhalter1 }},</p>
    <p>vielen Dank für Ihr Interesse an {{ Platzhalter2 }}.</p>
    <p>Mit freundlichen Grüßen,<br>Ihr Team</p>
  </body>
</html>
🔹 Hinweise
E-Mails werden über Outlook gesendet (Outlook muss installiert und konfiguriert sein).

Es wird kein externer Server benötigt.

Für den Testbetrieb kann -dryrun verwendet werden, dann werden die E-Mails nicht gesendet, sondern nur angezeigt.

yaml
Kopieren






