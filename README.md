# AbiKursliste2xlsx

Projekt um die aus Indiware exportierten Kurslisten nach Excel umzuwandeln.

- [AbiKursliste2xlsx](#abikursliste2xlsx)
  - [Projekt für Windows bauen](#projekt-für-windows-bauen)
  - [Kommandozeileargument](#kommandozeileargument)
  - [Umwandlung ausführen](#umwandlung-ausführen)
    - [Erstellung einer XLSX Datei mit mehreren Blättern](#erstellung-einer-xlsx-datei-mit-mehreren-blättern)
    - [Erstellung mehrerer XLSX Dateien](#erstellung-mehrerer-xlsx-dateien)
  - [sonstiges](#sonstiges)
    - [Umwandlung mit pdftotext nach text](#umwandlung-mit-pdftotext-nach-text)

## Projekt für Windows bauen

```shell
GOOS=windows GOARCH=amd64 go build
```

## Kommandozeileargument

```shell
-files
      Generate multiple XLSX files.
-pdf string
      Path to an PDF file.
-sheets
      Generate an XLSX file with multiple sheets. (default true)
```

## Umwandlung ausführen

### Erstellung einer XLSX Datei mit mehreren Blättern

```shell
./AbiKursliste2ods -pdf AbiKursliste.pdf
```
### Erstellung mehrerer XLSX Dateien

```shell
./AbiKursliste2ods -files -pdf AbiKursliste.pdf
```

## sonstiges

### Umwandlung mit pdftotext nach text

```shell
pdftotext -layout -nopgbrk AbiKursliste.pdf
```
