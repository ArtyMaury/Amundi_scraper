# Amundi_scraper

## Setup

1. Create a google spreadsheet
2. Go to tools -> Script editor
3. Paste the main.gs file
4. Login to amundi, get the "X-noee-authorization" header content and set it as API_TOKEN
5. Execute

6. Make a sheet with the data you want

## Interesting data example:

Operations:
dateDeLaDemande(R)	montantNet(AK)	montantNetAbondement(AL)	libelleCommunication(AB)

Dispositifs:
libelleDispositifMetier(AF)	mtBrut(AH)	mtNet(AK)	mtBrutDispo(AI)

## NB

The token has a very short timeout so it's not that sensitive
