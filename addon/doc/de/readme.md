# simpleCopy

**Texte kopieren, anhängen und effizient mit NVDA verwalten**

**Autor:** chai chaimee  
**URL:** https://github.com/chaichaimee/simpleCopy

---

## Beschreibung

**simpleCopy** ist ein leichtgewichtiges NVDA-Add-on, das das Kopieren von Texten, das Extrahieren von URLs und die Verwaltung des Sprachverlaufs vereinfacht.

Dieses Tool hilft Ihnen, Informationen schnell zu erfassen und zu organisieren, ohne Ihren Arbeitsablauf zu unterbrechen. Ob Sie Texte kopieren, Web-Links erfassen oder gesprochene Inhalte speichern möchten – simpleCopy bietet intuitive Tastaturkürzel, die nahtlos mit NVDA funktionieren.

---

## Tastenkürzel

Alle Befehle verwenden ein Multi-Tap-System. Drücken Sie die Tastenkombination einmal, zweimal oder dreimal in schneller Folge, um verschiedene Aktionen auszuführen.

### STRG+Umschalt+A — URL- und Link-Erfassung

- **Einmal drücken:** Kopiert die aktuelle Webseiten-URL.
- **Zweimal drücken:** Kopiert die Ziel-URL des fokussierten Hyperlinks.

### STRG+Umschalt+V — Text kopieren, anhängen und Zwischenablage verwalten

- **Einmal drücken:** Kopiert den ausgewählten Text. Wenn sich bereits Text in der Zwischenablage befindet, wird die neue Auswahl angehängt.
- **Zweimal drücken:** Kopiert Text von der aktuellen PrüfCursor-Position. Dies funktioniert mit jeder Auswahl, die mit dem PrüfCursor von NVDA erstellt wurde, einschließlich mehrzeiliger Auswahl und vollständiger Dokumentauswahl.
- **Dreimal drücken:** Löscht den gesamten Inhalt der Zwischenablage.

### F9 — Sprachausgabe erfassen und verwalten

- **Einmal drücken:** Kopiert die letzte Sprachausgabe von NVDA.
- **Zweimal drücken:** Hängt die letzte Sprachausgabe an den vorhandenen Zwischenablageinhalt an.
- **Dreimal drücken:** Kopiert die gesamte seit dem ersten F9-Druck gesammelte Sprachausgabe.

### Umschalt+F9 — Sprachverlauf navigieren

- **Einmal drücken:** Navigiert zum vorherigen Eintrag im Sprachverlauf.
- **Zweimal drücken:** Navigiert zum nächsten Eintrag im Sprachverlauf.
- **Dreimal drücken:** Öffnet die vollständige Sprachverlaufs-Logdatei.

---

## Funktionen

So funktioniert jede Funktion in der Praxis:

### 1. Webseiten-URL kopieren

Drücken Sie **STRG+Umschalt+A einmal** während Sie eine Website besuchen. Die aktuelle Seiten-URL wird in Ihre Zwischenablage kopiert. NVDA bestätigt dies, indem es die kopierte URL vorliest.

### 2. Hyperlink-URL extrahieren

Fokussieren Sie einen Link und drücken Sie **STRG+Umschalt+A zweimal**. Die Ziel-URL wird extrahiert und kopiert, ohne den Link zu öffnen.

### 3. Text kopieren und anhängen

Markieren Sie Text und drücken Sie **STRG+Umschalt+V einmal**. Wenn die Zwischenablage leer ist, wird der Text kopiert. Wenn die Zwischenablage bereits Text enthält, wird die neue Auswahl mit einem Zeilenumbruch angehängt.

### 4. PrüfCursor kopieren

Verwenden Sie den PrüfCursor von NVDA, um Text auszuwählen (mit NVDA+Umschalt+Nach-unten oder NVDA+STRG+Umschalt+Nach-unten, um mehrere Zeilen auszuwählen), und drücken Sie dann **STRG+Umschalt+V zweimal**. Der gesamte ausgewählte Text von der PrüfCursor-Position wird in die Zwischenablage kopiert. Dies funktioniert mit jeder Auswahlgröße, von einem einzelnen Wort bis zu einem gesamten Dokument.

### 5. Zwischenablage leeren

Drücken Sie **STRG+Umschalt+V dreimal**, um den gesamten Inhalt der Zwischenablage sofort zu löschen. NVDA bestätigt dies mit der Meldung "Clean".

### 6. Letzte Sprachausgabe kopieren

Wenn NVDA etwas spricht, das Sie speichern möchten, drücken Sie **F9 einmal**. Die letzte gesprochene Phrase wird in Ihre Zwischenablage kopiert.

### 7. Sprachausgabe anhängen

Drücken Sie **F9 zweimal**, um die letzte gesprochene Phrase an den vorhandenen Zwischenablageinhalt anzuhängen.

### 8. Sprachverlauf protokollieren

Drücken Sie **F9 dreimal**, um die gesamte während Ihrer aktuellen Sitzung gesammelte Sprachausgabe zu kopieren.

### 9. Sprachverlauf navigieren

Verwenden Sie **Umschalt+F9 einmal**, um im Sprachverlauf rückwärts zu navigieren, und **Umschalt+F9 zweimal**, um vorwärts zu navigieren. So können Sie vergangene Sprachausgaben überprüfen, ohne Ihren aktuellen Fokus zu ändern.

### 10. Sprach-Logdatei öffnen

Drücken Sie **Umschalt+F9 dreimal**, um die vollständige Sprachverlaufsdatei in Ihrem Standard-Texteditor zu öffnen, um sie zu durchsuchen, zu bearbeiten oder zu kopieren.

### 11. Intelligente Kontexterkennung

Wenn Sie in bearbeitbaren Feldern tippen, stört simpleCopy nicht. Befehle werden nur aktiviert, wenn sie nützlich sind, und bewahren so Ihren normalen Arbeitsablauf.

---

## Unterstützen Sie mich

Wenn dieses Add-on Ihnen hilft, effizienter zu arbeiten, ziehen Sie bitte eine kleine Spende zur Unterstützung der zukünftigen Entwicklung in Betracht.

[![Unterstützen Sie mich](https://img.shields.io/badge/Donate-Support%20Me-blue?style=for-the-badge&logo=stripe)](https://buy.stripe.com/dRm9AU1xQ3Ds22N6VK1VK01)

Ihre Unterstützung hilft, dieses Projekt am Leben zu erhalten und zu verbessern.

---

© 2026 Chai Chaimee NVDA-Add-on Veröffentlicht unter der GNU General Public License