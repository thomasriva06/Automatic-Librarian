## Guida passo dopo passo (Download, installazione e utilizzo)

Di seguito spiego **come scaricare, installare e usare** il progetto **passo dopo passo**, con una procedura **step by step** pensata anche per chi non ha mai utilizzato Python o GitHub.  
L’obiettivo è farti arrivare rapidamente a un sistema funzionante, così da poter inserire i libri e ottenere in automatico il file Excel ordinato.

## Procedura in comune con macOS e Windows

**Formato con cui inserirai i libri (1 riga = 1 libro)**

- Separatore **;**

  esempio:
  ```text
  Cognome;Nome;Titolo;Collana;Casa editrice;Anno;Genere/Contesto
  
  ```
  esempio:
  ```text
  Virgilio;Publio Marone;Eneide;Classici Latini;Mondadori;1985;Letteratura latina
  ```

## Dove tenere i file

Metti **insieme** (nella stessa cartella):

	•	biblioteca.xlsx    --> lo scarichiamo insieme nei prossimi step
	•	biblioteca_manager.py    --> scaricabile dai file caricati su questo GitHub

E lavora sempre lì.

## Procedura per Mac (macOS)
**Step 1 — Crea la cartella di lavoro**
1.	Crea una cartella, ad esempio: Biblioteca
2.	Metti dentro *biblioteca.xlsx* e *biblioteca_manager.py*

**Step 2 — Installa Python (se non gia installato)**
1.	Apri **Terminale** (Applicazioni → Utility → Terminale)
2.	Scrivi:
```text
python3 --version
```
- Se ti mostra una versione (es. 3.11.x), ok.
- Se non va, installa Python 3 da *python.org*.

## Step 3 — Installa la libreria necessaria

Nel Terminale, entra nella cartella dove hai i file. Esempio se è in Scrivania:

```text
cd ~/Desktop/Biblioteca
```
Poi installa:
```text
python3 -m pip install --upgrade pip
python3 -m pip install openpyxl
```
## Step 4 — Biblioteca.xlsx

Questo crea il file Excel che dicevo precedentemente, il quale va inserito all'interno della cartella che abbiamo creato poco fa (Biblioteca):
```text
python3 biblioteca_manager.py init --file biblioteca.xlsx
```
## Step 5 — Aggiungi libri (incollando liste)

Esegui:
```text
python3 biblioteca_manager.py add --file biblioteca.xlsx
```
Poi:
1.	Incolla molte righe (una per libro)
2.	Quando hai finito, premi **Ctrl + D** (chiude l’input)

Risultato: *biblioteca.xlsx* viene aggiornato e riordinato.

## Step 6 — Conclusione

Ogni volta che vuoi aggiungere nuovi libri, ad esempio nei giorni successivi, basterà incollare la seguente stringa di codice:
```text
cd ~/Desktop/Biblioteca
python3 biblioteca_manager.py add --file biblioteca.xlsx
```

***AVVISO:*** se ti da errore significa che stai lanciando lo script da una cartella diversa da quella in cui si trova *biblioteca_manager.py*, perciò basta semplicemente cambiare **cd ~/Desktop/Biblioteca** con la posizione corretta.

## Procedura per Windows (Windows 10/11)
**Step 1 — Crea la cartella di lavoro**
1.	Crea una cartella, ad esempio: *C:\Biblioteca*
2.	Metti dentro *biblioteca.xlsx* e *biblioteca_manager.py*

**Step 2 — Installa Python (se non gia installato)**
1.	Apri **Windows Terminal** oppure **Prompt** dei comandi
2.	Digita:
```text
python --version
```
- Se esce una versione, ok.
- Se non va: installa Python 3 da python.org e durante l’installazione spunta ***“Add Python to PATH”.***

## Step 3 — Installa la libreria necessaria
1.	Vai nella cartella di lavoro:
```text
cd C:\Biblioteca
```
2.	Installa:

```text
python -m pip install --upgrade pip
python -m pip install openpyxl
```
## Step 4 — Biblioteca.xlsx

Questo crea il file Excel che dicevo precedentemente, il quale va inserito all'interno della cartella che abbiamo creato poco fa (Biblioteca):
```text
python biblioteca_manager.py init --file biblioteca.xlsx
```
## Step 5 — Aggiungi libri (incollando liste)

Esegui

```text
python biblioteca_manager.py add --file biblioteca.xlsx
```
Poi:
1.	Incolla molte righe
2.	Quando hai finito:
  -	Premi ***Ctrl + Z***
  -	Poi ***Invio***

Risultato: Excel aggiornato e riordinato.

## Step 6 — Conclusione

Ogni volta che vuoi aggiungere nuovi libri, ad esempio nei giorni successivi, basterà incollare la seguente stringa di codice:

```text
cd C:\Biblioteca
python biblioteca_manager.py add --file biblioteca.xlsx
```
***AVVISO:*** se ti da errore significa che stai lanciando lo script da una cartella diversa da quella in cui si trova *biblioteca_manager.py*, perciò basta semplicemente cambiare ***cd C:\Biblioteca*** con la posizione corretta.
