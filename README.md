# Simulatore della bolletta della luce

Calcolatore della spesa annua per l'elettricità di una famiglia, con e senza il tetto al prezzo del gas previsto dal Decreto Bollette bis (Consiglio dei ministri del 9 novembre 2026).

Pagina pubblicata: https://liberostracquadanio-del.github.io/simulatore-bolletta-luce/

## Metodo
- Tariffe: ARERA, condizioni economiche per i clienti vulnerabili in maggior tutela, ottobre-dicembre 2026 (https://www.arera.it/fileadmin//area_operatori/prezzi_e_tariffe/E2026-smt.xlsx), applicate a dodici mesi.
- Imposte: accisa di 0,0227 euro/kWh con l'esenzione per le abitazioni di residenza fino a 3 kW (D.Lgs. 504/1995, art. 52); IVA al 10%.
- Decreto: taglio di 15 euro/MWh sul prezzo all'ingrosso (PUN), maggiorato del 10% per le perdite di rete in bassa tensione (ARERA, Tabella 4 del TIS), e onere pari al 40% del beneficio. Sono stime del ministro Pichetto Fratin: il testo del decreto non era ancora in Gazzetta Ufficiale al 10 novembre 2026.
- Controllo: con 2.000 kWh e 3 kW il calcolo restituisce 43,43 centesimi/kWh, come il cliente tipo del comunicato ARERA del 29 settembre 2026.

## Incorporamento
```html
<iframe src="https://liberostracquadanio-del.github.io/simulatore-bolletta-luce/" title="Simulatore della bolletta della luce" style="width:100%;border:0;height:1000px" loading="lazy"></iframe>
```
Il file `index.html` è generato da uno script a partire dai dati ARERA: non va modificato a mano.
