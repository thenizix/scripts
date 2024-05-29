D.Nencini 5/2024
Questo script nasce dall'esigenza marginale di convertire documenti word in tabelle di excel.
Nella speranza di simulare un database efficiente.
Seppure contrario all obsoleto percorso, mi ritrovo con la sgradevole situazione di avere delle macchine della PA.
Senza privilegi da amministratore, con windows 10 non aggiornato e senza wsl.
Non posso ovviamente eseguire nessun programma, ma ho voluto lo stesso mettere insieme qualche riga di codice e ....
funziona

Per eseguire il programma da windows wsl, scaricate il codice   script_conversione.py
Mettete uno o cento files .docx nella stessa directory dello script.
Aprite una shell e scrivete 
            python script_conversione.py
il programma dovrebbe installare automaticamente le dipendenze
ed eseguirsi.
Se i files sono quelli giusti la conversione diventa velocissima.

-----------------------------------
questo script e' solo un test la definizione dei record e' approssimata e non customizzabile al volo
Servirebbe un interfaccia grafica ma da wsl la vedo male.
 ho provato con altri linguaggi 
ma non sono riuscito a creare un applicazione che non necessitasse di accesso admin.
Alla fine ho scelto Python.

