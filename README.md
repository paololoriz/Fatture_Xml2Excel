<h1>FattureXml2Excel.exe<h1>
<h3>Premessa<h3>
Per l’utilizzo dello script è necessario disporre in una cartella tutte le fatture elettroniche in formato xml. 
Inserire file con un formato differente da quello xml potrebbe determinare errori nell’esecuzione dello script.
È necessario organizzare il file FattureXml2Excel in una cartella in quanto il foglio Excel che verrà generato verrà generato nella stessa cartella di origine del file dello script.
<h3>Apertura<h3>
All’apertura del file .exe FattureXml2Excel verrà aperta una finestra di dialogo per scegliere la cartella di origine (con i file delle fatture elettroniche).
<h3>Conclusione<h3>
Dopo aver scelto la cartella lo script procederà in modo autonomo e genererà il file Excel.
Il file è organizzato in modo tale che ogni riga rappresenti una fattura con tutti i dati utili ad essa collegati.
La colonna A (che rappresenta il numero della fattura) potrebbe essere colorata in:
•	arancione quando la fattura presenta più modalità di pagamento e non è possibile associare la giusta modalità al pagamento
•	 rosso quando la fattura presenta dati pagamento mancanti (es. caso di uno storno)
