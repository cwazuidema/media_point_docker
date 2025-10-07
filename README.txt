# Media Point Excel Processor (Dockerversie)

Verwerk `Bron.xlsx` automatisch naar `Modified_Bron.xlsx` via een eenvoudige webinterface in Docker.

## Vereisten

- Windows 10/11 met Docker Desktop actief
- Internetverbinding voor de eerste build

## Snelstart (Windows)

1. Dubbelklik op `start.bat` in deze map.
2. De browser opent `http://localhost:8000`.
3. Op de pagina:
   - Selecteer `Bron.xlsx` en klik op "Upload Bron.xlsx".
   - Wacht tot de melding "Download is ready." verschijnt.
   - Klik op "Download Modified_Bron.xlsx" om het resultaat op te slaan.

De bestanden worden direct in deze map gelezen en weggeschreven via de Docker-volume koppeling (`./:/data`).

## Werking

- Een kleine FastAPI-webserver draait in de container (poort 8000).
- De bestaande `lambda_handler` in `excel_processor.py` verwerkt lokale bestanden:
  - Input: `/data/Bron.xlsx` (via de volume mount)
  - Output: `/data/Modified_Bron.xlsx`
- FTP-functionaliteit is verwijderd.

## Projectstructuur

- `app/main.py`: FastAPI-applicatie met upload- en downloadlogica
- `excel_processor.py`: verwerkingslogica op basis van de oorspronkelijke Lambda-code
- `Dockerfile`, `docker-compose.yml`, `.dockerignore`, `requirements.txt`
- `start.bat`: snelle start voor Windows-gebruikers

## Opmerkingen

- Behoud domeinspecifieke logica (bijv. Nederlandstalige kolomnamen) indien aanwezig.
- Bij een grote `Bron.xlsx` kan de eerste build langer duren wegens het downloaden van Python-wheels.
