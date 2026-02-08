<div align="center">

# 🏗️ NP Zero
### Gestionale Analisi Nuovi Prezzi per l'Edilizia

![Python](https://img.shields.io/badge/Python-3.x-blue?style=for-the-badge&logo=python)
![GUI](https://img.shields.io/badge/Interface-Tkinter-orange?style=for-the-badge)
![License](https://img.shields.io/badge/License-GPLv3-green?style=for-the-badge)
![Status](https://img.shields.io/badge/Status-v1.1.0_Multi--User-purple?style=for-the-badge)

**Basta fogli di calcolo complessi e analisi disorganizzate.** NP Zero è lo strumento open-source definitivo per tecnici e professionisti dell'edilizia. Crea, calcola e gestisci i tuoi Nuovi Prezzi (NP) in modo professionale, con generazione automatica di analisi prezzi in PDF.

[Caratteristiche](#-caratteristiche-principali) • [Installazione](#-installazione) • [Workflow](#-come-usare) • [Struttura](#-struttura-del-progetto)

</div>

---

## 🧐 Cos'è NP Zero?

**NP Zero** nasce per semplificare uno dei compiti più delicati per un progettista o un direttore dei lavori: la determinazione di prezzi non presenti nei prezzari ufficiali (analisi dei nuovi prezzi).

Ti aiuta a rispondere alle sfide quotidiane del cantiere:
* *"Come scompongo correttamente questa lavorazione tra manodopera e materiali?"*
* *"Qual è l'incidenza reale della sicurezza all'interno delle spese generali?"*
* *"Come posso generare una stampa professionale per la stazione appaltante?"*

Questo software automatizza il calcolo dei costi diretti e indiretti, gestisce l'anagrafica dei progetti e ti permette di riutilizzare analisi esistenti in pochi clic.

## ✨ Caratteristiche Principali

* 📁 **Gestione Progetti:** Organizza i tuoi NP per intervento, associando Codice, Titolo, CUP e Committente.
* 🧮 **Analisi Dettagliata:** Scomposizione analitica dei costi in tre categorie standard: **Manodopera**, **Mezzi e Noli**, **Materiali**.
* 📈 **Calcolo Automatico SG & Utili:** Gestione dinamica delle Spese Generali (default 17%) e Utili d'Impresa (default 10%), con ricalcolo immediato del prezzo finale.
* 📄 **Esportazione HTML e PDF:** Generazione di analisi prezzi professionali in formato HTML, convertibili massivamente in PDF (singoli o uniti) pronti per l'allegato tecnico.
* 🔄 **Smart Import:** Funzione per importare e duplicare NP da altri progetti, risparmiando tempo sulle lavorazioni ricorrenti.
* 📊 **Riepilogo Incidenze:** Calcolo automatico delle incidenze percentuali di manodopera e sicurezza, fondamentali per la redazione dei progetti.
* 💾 **Database Locale:** Massima privacy e portabilità. Tutti i dati sono salvati in un database SQLite locale sul tuo PC.

## 🚀 Installazione

Puoi utilizzare l'eseguibile compilato tramite GitHub Actions (se configurato) o eseguire direttamente i sorgenti Python.

### Prerequisiti
* Python 3.10 o superiore.
* Un browser basato su Chromium (Google Chrome o Microsoft Edge) installato per la conversione PDF.

### Passaggi

1.  **Clona il repository** (o scarica lo zip):
    ```bash
    git clone [https://github.com/piano-zero/np-zero.git](https://github.com/piano-zero/np-zero.git)
    cd np-zero
    ```

2.  **Installa le dipendenze:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Avvia l'applicazione:**
    ```bash
    python Rpo_Zero_v2.0.0.py
    ```

## 🛠 Struttura del Progetto

Il progetto è strutturato per essere leggero e privo di configurazioni esterne:

* `Rpo_Zero_v2.0.0.py` 🧠: Il cuore dell'applicazione. Contiene la logica della GUI (Tkinter), la gestione del database SQLite e il motore di generazione stampe.
* `requirements.txt` 📋: Elenco delle librerie necessarie (come `pypdf` per la manipolazione dei file).
* `np_zero.db` 💾: Database locale creato automaticamente al primo avvio.
* `NP_STAMPE/` 📂: Cartella generata automaticamente dove vengono archiviati i file HTML e PDF esportati.

## 📖 Come Usare (Workflow)

1.  **Crea Progetto:** Inizia dalla Tab "Progetti" inserendo i dati del tuo intervento o cantiere.
2.  **Definisci NP:** Nella Tab "Elenco NP", crea una nuova voce di prezzo associata al progetto selezionato.
3.  **Dettaglio Costi:** Entra in "Modifica NP" per inserire le singole voci (ore uomo, materiali, noli). Il software sommerà tutto e aggiungerà automaticamente le aliquote SG e Utili.
4.  **Esporta:** Vai nella Tab "Stampe" per generare l'HTML e successivamente nella Tab "Convertitore PDF" per ottenere i documenti finali.

## 🤝 Contribuire

Le idee per migliorare NP Zero sono sempre benvenute! Se vuoi aggiungere una funzionalità (es. esportazione Excel):

1.  Fai un **Fork** del progetto.
2.  Crea un branch per la tua modifica (`git checkout -b feature/MiglioriaTecnica`).
3.  Fai **Commit** (`git commit -m 'Aggiunta nuova funzione'`).
4.  Fai **Push** (`git push origin feature/MiglioriaTecnica`).
5.  Apri una **Pull Request**.

## 📄 Licenza

Distribuito sotto licenza **MIT**.

---

<div align="center">
  
  Created with ❤️ by [pianozero](https://github.com/piano-zero)
  
  *Se questo progetto ti aiuta nel tuo lavoro tecnico, lascia una ⭐️ al repository!*

</div>
