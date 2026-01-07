# Automatic-Librarian
Library reorganization tool: from random input, create/update an Excel spreadsheet with books grouped by context and sorted by author, year, and volume, ready for printing or consultation.


## Biblioteca Excel Manager

Sono **Riva Thomas**, studente di **Ingegneria Informatica** al **Politecnico di Milano**.  
Ho realizzato questo progetto in **Python** per automatizzare la catalogazione e l’ordinamento di **grandissime quantità di libri** in modo rapido e affidabile.

## Perché nasce questo progetto

Il progetto nasce da un’esigenza concreta: **riordinare e ricatalogare la biblioteca scolastica** della mia scuola superiore.  
Dato il **poco tempo a disposizione** e l’elevato numero di volumi da inserire manualmente, ho scelto la strada più efficiente: **automatizzare il processo**.

L’idea è semplice: invece di compilare a mano un catalogo, una persona può **dettare vocalmente** le informazioni dei libri (tramite **Whisper** o altri sistemi di dettatura).  
Il software di dettatura trascrive il testo e, una volta ottenuti i dati in formato strutturato, il mio programma li elabora automaticamente.

## Come nasce l’input e come si usa davvero (in modo pratico)

Dietro questo progetto non c’è solo “un programma che scrive su Excel”, ma un’idea pensata per rendere **realmente fattibile** un lavoro che, fatto a mano, richiederebbe giorni (se non settimane): ricatalogare un’intera biblioteca con molti volumi, spesso diversi tra loro per formato e provenienza.

La parte più delicata, infatti, non è tanto “salvare i dati”, ma **acquisirli velocemente e in modo uniforme**. Per questo ho scelto un formato d’inserimento semplice e ripetibile:

**`Cognome;Nome;Titolo;Collana;Casa editrice;Anno;Genere/Contesto`**

Non è una scelta casuale: la maggior parte dei libri riporta queste informazioni in maniera molto simile **sulla copertina o nel frontespizio**, e spesso, leggendo dall’alto verso il basso, l’ordine è proprio questo. Mantenere la stessa sequenza rende il lavoro più rapido e riduce errori, soprattutto quando si sta trascrivendo a voce con strumenti come **Whisper** o altri sistemi di dettatura: chi detta (o chi controlla la trascrizione) sa già esattamente quale campo sta inserendo in ogni posizione.

### Inserimento “a blocchi” (una riga o molte righe insieme)

Puoi inserire:
- **una sola riga** (un libro), oppure
- **più righe tutte insieme** (molti libri in batch)

Ogni riga deve contenere **tutti i campi**, separati da **punto e virgola**.  
Quando hai completato i dati di un libro, **vai a capo** e scrivi il libro successivo, sempre nello stesso formato.

Esempio (più libri incollati in un colpo solo):

```text
Virgilio;Publio Marone;Eneide;Classici Latini;Einaudi;19;Letteratura latina
Austen;Jane;Pride and Prejudice;Classics;Penguin;1813;Letteratura inglese
Hugo;Victor;Les Misérables;Classiques;Gallimard;1862;Letteratura francese
```

## Cosa fa il programma

Il programma:
- **crea** (se non esiste) un file **Excel**
- **aggiunge** nuovi libri in modo incrementale, anche su più giorni
- **aggiorna e riordina** continuamente il catalogo in base a regole precise e ripetibili
- permette di lavorare “a blocchi”: inserisci molti libri alla volta, e il file viene sistemato automaticamente

## Input richiesto

Il programma lavora a partire dai campi (nell’ordine):

- **Cognome**
- **Nome**
- **Titolo**
- **Collana**
- **Casa editrice**
- **Anno di pubblicazione**
- **Contesto di appartenenza** (letteratura inglese, letteratura francese, letteratura latina,...)

Una volta raccolti i dati (anche tramite dettatura + trascrizione), essi vengono inseriti nel programma in questione e il resto avviene in modo automatico.

## Criterio di ordinamento

I libri vengono organizzati nel file Excel seguendo esattamente questo criterio:

1. **Genere/Contesto** (es. *Letteratura latina, greca, inglese, francese, tedesca…*)
2. All’interno di ogni contesto: **Cognome autore** in ordine alfabetico (A → Z)
3. Poi **Nome autore** (A → Z)
4. Per ogni autore e contesto: **Anno di pubblicazione** dal più vecchio al più recente
5. Se due libri hanno **stesso autore + stesso contesto + stesso anno**, il programma usa l’ordine dei **volumi**:
   - *Volume primo, secondo, terzo, quarto, …* (o equivalenti come *Vol. I, II, III…*)

## Flusso di lavoro tipico

1. Dettatura dei libri (Whisper o altro sistema)  
2. Trascrizione testuale dei campi richiesti 
3. Inserimento dei record nel programma (anche in batch)  
4. Generazione/aggiornamento dell’Excel con ordinamento automatico

## Obiettivo

Ridurre drasticamente il tempo di catalogazione manuale e ottenere un **catalogo ordinato, coerente e sempre aggiornato**, anche quando il lavoro si svolge su più giornate.
