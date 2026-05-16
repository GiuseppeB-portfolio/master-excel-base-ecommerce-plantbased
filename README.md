# Analisi delle vendite di un e-commerce plant-based

> Progetto di Excel Base realizzato nell'ambito del Master Universitario in AI e Agenti AI per il Business (Unimarconi · start2impact).

![status](https://img.shields.io/badge/status-completato-success)
![modulo](https://img.shields.io/badge/modulo-Analisi%20Dati%20e%20AI-blue)
![tool](https://img.shields.io/badge/tool-Microsoft%20Excel-217346?logo=microsoftexcel&logoColor=white)

---

## 📋 In sintesi

|  |  |
|---|---|
| **Corso** | Master in AI e Agenti AI per il Business |
| **Modulo** | 1 — Analisi dei Dati e AI (Excel Base) |
| **Strumenti** | Microsoft Excel — funzioni, formattazione condizionale, tabelle pivot, grafici |
| **Cliente (ipotetico)** | POF — Planty of Food, e-commerce alimentare plant-based e sostenibile |
| **Dataset** | 25 scontrini distribuiti su province del Lazio e mesi dell'anno |
| **SDG di riferimento** | Consumo e produzione responsabili · Salute e benessere |

---

## 🎯 Contesto

**POF — Planty of Food** è un cliente ipotetico ma realistico: un e-commerce italiano specializzato in alimentazione plant-based, con produttori biologici selezionati, packaging compostabili e una formula di Gruppi di Acquisto Solidale (GAS) per ridurre costi di spedizione ed emissioni. Nato come pioniere del settore, oggi punta a diventarne il leader.

Per supportare le decisioni commerciali e di marketing, l'azienda ha bisogno di analizzare i propri dati di vendita per individuare i reparti più performanti, la stagionalità degli acquisti e la distribuzione geografica della clientela.

Il progetto risponde a una traccia strutturata di **20 quesiti** che richiedono l'uso progressivo di funzioni, formule condizionali, tabelle pivot e grafici.

---

## 📊 Il dataset

Il foglio principale `Food E-commerce Plant Based` contiene 25 scontrini con i seguenti campi:

- **# Scontrino** — identificativo della transazione
- **Provincia** — provincia di vendita (Roma, Latina, Frosinone, Viterbo, Rieti)
- **Mese** — mese della transazione
- **Importo Scontrino** — valore in euro
- **Reparto** — categoria di prodotto (Dispensa, Banco frigo, Bevande, Cereali, Ortofrutticolo, Vitamine e integratori)

Un secondo foglio `Città` contiene la mappatura provincia → città capoluogo, usata per arricchire il dataset principale tramite ricerca verticale.

Sono state aggiunte due colonne calcolate:

| Campo | Formula | Significato |
|---|---|---|
| Città | `=VLOOKUP(B2;Città!A:B;2;FALSO)` | Popolamento automatico della città a partire dalla provincia |
| Importo (fascia) | `=SE(D2>63,91;"Alto";"Basso")` | Classificazione dello scontrino rispetto alla media |

---

## 🧩 I 20 quesiti e le funzioni applicate

La traccia copre progressivamente diversi livelli di complessità delle formule Excel.

### Funzioni di aggregazione e filtro condizionale
| # | Quesito | Funzione |
|---|---|---|
| 5 | Totale vendite di tutti i mesi | `SOMMA` |
| 6 | Totale vendite del mese di giugno | `SOMMA.SE` |
| 9 | Conteggio degli scontrini sopra i 100 € | `CONTA.SE` |
| 10 | Totale vendite della provincia di Latina | `SOMMA.SE` |
| 11 | Numero di vendite a Roma ad aprile | `CONTA.PIÙ.SE` |
| 12 | Vendite da Dispensa nel mese di gennaio | `SOMMA.PIÙ.SE` |
| 13 | Importo medio di uno scontrino | `MEDIA` |
| 14–15 | Importo medio con filtro su soglie (> 20 €, < 106 €) | `MEDIA.SE` |

### Manipolazione e arricchimento dati
| # | Quesito | Funzione |
|---|---|---|
| 7 | Ordinamento decrescente per importo | Strumenti di ordinamento |
| 8 | Evidenziare gli importi > 100 € | Formattazione condizionale |
| 16 | Popolare la colonna Città dal secondo foglio | `CERCA.VERT` |
| 17 | Etichetta "Alto" / "Basso" rispetto a una soglia | `SE` |

### Visualizzazione e sintesi
| # | Quesito | Soluzione |
|---|---|---|
| 18 | Tabella Pivot mesi × reparti × importo | Pivot su foglio dedicato + `INDICE/CONFRONTA` per individuare il reparto top |
| 19 | Grafico delle vendite mensili | Foglio dedicato + grafico a colonne |
| 20 | Mese con più vendite | `INDICE/CONFRONTA` sul foglio Grafico |

L'uso di `INDICE/CONFRONTA` sui quesiti 18 e 20 rende le risposte dinamiche: anche cambiando i dati di partenza, la cella risposta si aggiorna automaticamente con il nuovo reparto e il nuovo mese top, senza bisogno di leggere a occhio la pivot o il grafico.

---

## 💡 Insight emersi dall'analisi

Dai dati elaborati emergono alcune indicazioni operative interessanti per il cliente.

**Aprile è il mese con il maggior volume di vendite** (327,25 €), seguito a distanza da gennaio (288,89 €). Esistono mesi con attività molto bassa — luglio in particolare (35,71 €) — che potrebbero suggerire l'opportunità di campagne stagionali mirate.

**Banco frigo è il reparto più venduto** in termini di numero di transazioni (6 scontrini su 25), seguito da Dispensa e Vitamine e integratori (5 ciascuno). I dati indicano un mix di acquisto orientato sia ai prodotti freschi sia ai supplementi nutrizionali — coerente con il posizionamento plant-based del brand.

**La distribuzione geografica è concentrata nel Lazio**: la base clienti è raggiungibile via GAS in modo efficiente, ma anche un segnale che il bacino è ancora regionale e c'è ampio margine di espansione su scala nazionale.

---

## 📁 Struttura del file

```
Progetto_Excel_di_Giuseppe_Bianco.xlsx
├── Food E-commerce Plant Based   → foglio principale con dataset, formule e risposte
├── Città                          → mappatura provincia → città capoluogo
├── Tabella Pivot                  → pivot mesi × reparti × importo
└── Grafico                        → dati aggregati per mese e grafico a colonne
```

---

## ✍️ Riflessione finale

> Questo corso mi ha permesso di sviluppare un metodo più strutturato ed efficiente per analizzare i dati: impostare fogli di lavoro più chiari, applicare formule e funzioni in modo consapevole, presentare le informazioni rendendole comprensibili a colpo d'occhio. Le competenze acquisite sono immediatamente applicabili nel mio lavoro quotidiano, soprattutto quando devo pulire e organizzare dati in modo rapido, automatizzare controlli e calcoli con `SE`, `SOMMA.SE`, `CERCA.VERT`, e rendere i report più leggibili con formattazioni coerenti e strutture ordinate.

---

## 🚀 Come consultare il progetto

1. Scarica il file `Progetto_Excel_di_Giuseppe_Bianco.xlsx`
2. Aprilo con Microsoft Excel (o LibreOffice Calc / Google Sheets)
3. Le risposte ai quesiti si trovano nelle colonne H–J del foglio principale, accanto al testo della domanda

---

*Progetto realizzato nell'ambito del Master Universitario in AI e Agenti AI per il Business — Università degli Studi Guglielmo Marconi e start2impact.*
