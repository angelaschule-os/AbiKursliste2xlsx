# AbiKursliste2xlsx

Projekt um die aus Indiware exportierten Kurslisten nach Excel umzuwandeln.

## Projekt für Windows bauen

```shell
GOOS=windows GOARCH=amd64 go build
```

## Umwandlung ausführen

```shell
./AbiKursliste2ods AbiKursliste.pdf
```

## sonstiges

### Umwandlung mit pdftotext nach text

```shell
pdftotext -layout -nopgbrk AbiKursliste.pdf
```
