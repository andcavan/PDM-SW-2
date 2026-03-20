# PDM-SW – Sistema PDM per SolidWorks

> **Versione 2.10.20** - Modulo Commerciali/Normalizzati: categoria per tipo, codici manuali, duplica articolo, percorso archivio configurabile

Sistema PDM (Product Data Management) leggero, senza server, per la gestione di documenti SolidWorks in ambiente di rete con architettura peer‑to‑peer.

---

## Installazione su nuovo PC

1. Copia la cartella `PDM-SW-2/` in qualsiasi percorso (es. `C:\Lavoro\PDM-SW-2\`)
2. Esegui **`install.bat`** — verifica Python, crea `.venv`, installa dipendenze
3. Esegui **`start.bat`** — al primo avvio appare il wizard di configurazione:
   - Scegli la **cartella dati locali** (es. `C:\Users\mario\PDM-Data\`) — conterrà `local_config.json`
   - Scegli la **workspace SolidWorks** — cartella locale per i file in checkout
4. Il wizard guida poi alla configurazione del percorso condiviso di rete e al login

> **Nota**: `.venv/` e `local_config.json` **non** vengono copiati nella distribuzione. Ogni PC genera i propri.

## Creazione distribuzione

```
python make_dist.py           # crea dist/PDM-SW-2/
python make_dist.py --zip     # crea anche dist/PDM-SW-2_v2.10.19.zip
python make_dist.py --clean   # pulisce dist/ prima di procedere
```

---

## Changelog

### v2.10.20 — 2026-03-20
**Modulo Commerciali/Normalizzati: miglioramenti e nuove funzionalità**

- **`core/database.py`** — aggiunta colonna `item_type` in `commercial_categories` (con migrazione ALTER TABLE); constraint UNIQUE cambiato in `(item_type, code)`; rimossa colonna `active`.
- **`core/commercial_coding_config.py`** — codici categoria/sottocategoria a 4 cifre; aggiunto campo `commercial_archive_path`; metodi helper `next_cat_code`, `next_sub_code`.
- **`core/commercial_manager.py`** — `get_categories` e `create_category` filtrano/accettano `item_type`; aggiunto `_commercial_archive_root()` per percorso archivio configurabile; aggiornati `_archive_file_path()` e `link_sw_file()`; aggiunto `duplicate_item()`.
- **`ui/commercial_category_dialog.py`** — riscrittura completa: due schede (Commerciali/Normalizzati), box separati categorie e sottocategorie, campi codice editabili con pre-compilazione suggerita, rimosso campo "Attivo".
- **`ui/commercial_view.py`** — colonne albero ampliate a 7 (Descrizione, Sottocategoria); pulsante "Nuovo" spostato in toolbar; rimosso pulsante ↻ (usa toolbar); aggiunta voce "Duplica articolo" nel menu contestuale.
- **`ui/commercial_item_dialog.py`** — sottocategoria obbligatoria; supporto parametro `prefill` per duplica (mostra finestra pre-compilata prima di generare codice); filtro categorie per tipo articolo.
- **`ui/commercial_settings_dialog.py`** — nuovo dialog per configurare il percorso archivio file SW commerciali.
- **`ui/main_window.py`** — aggiunto pulsante "🏷️ Nuovo commerciale" in toolbar; aggiunta voce "⚙️ Impostazioni commerciali…" nel menu Strumenti.
- **`config.py`** — versione aggiornata a `2.10.20`.

### v2.10.19
**Fix: alias duplicati e mancato aggiornamento title/description**

- **`core/properties_manager.py`** - `_pick_first_prop_value` ora ignora i match vuoti e continua sul prossimo alias disponibile.
- **`core/properties_manager.py`** - nuova normalizzazione alias in import (`Title/Titolo`, `Description/Descrizione`, `Author/Autore`) per evitare doppioni in `document_properties`.
- **`core/properties_manager.py`** - default mapping `title` e `description` impostati su `Bidirezionale` (oltre a `revision`) per rendere effettivo il sync PDM->SW anche su questi campi.
- **`ui/sw_config_dialog.py`** - default UI allineati su `Bidirezionale` per `Titolo` e `Descrizione`.
- **`config.py`** - versione applicazione aggiornata a `2.10.19`.

### v2.10.18
**Fix proprieta non movimentate (SW<->PDM)**

- **`core/properties_manager.py`** - `_read_custom_props` reso compatibile con varianti COM dove `Count`/`GetNames` sono metodi o tuple; evita salti silenziosi che portavano a importazioni vuote.
- **`core/properties_manager.py`** - `write_to_sw_file` migliorato con fallback su documento aperto per nome file, scrittura `Set2`/`Add3`/`Add2` e salvataggio `Save3` con fallback `Save`.
- **`macros/sw_bridge.py`** - `action_export_props` ora usa la sync mapping `sync_pdm_to_sw` (campi PDM fondamentali) e poi esporta le custom properties.
- **`config.py`** - versione applicazione aggiornata a `2.10.18`.

### v2.10.17
**Fix sincronizzazione proprieta e aggiornamento campi PDM**

- **`core/properties_manager.py`** - matching nomi proprieta reso robusto (ignora spazi/simboli) per intercettare varianti reali dei nomi SW.
- **`core/properties_manager.py`** - lettura SW arricchita con `SummaryInfo` (`Title`, `Subject`, `Comments`, `Author`, `CreatedDate`, `SaveDate`) e fallback `Description` da `Subject/Comments`.
- **`ui/document_dialog.py`** - `Importa da SW` ora ricarica sempre il documento e segnala correttamente quando il mapping aggiorna `title/description` anche senza custom properties.
- **`macros/sw_bridge.py`** - `import_props` non segnala piu errore se non ci sono custom properties ma il mapping ha aggiornato i campi PDM.
- **`ui/sw_config_dialog.py`** - alias default mapping allineati ai nuovi fallback (`Subject`, `Comments`, `Author`, `Date`).
- **`config.py`** - versione applicazione aggiornata a `2.10.17`.

### v2.10.16
**Mappature proprieta estese: codice, stato, creato da, data**

- **`ui/sw_config_dialog.py`** - tab **`Mappature proprieta`** estesa con i campi PDM `code`, `state`, `created_by`, `created_at`, mostrando esplicitamente il nome proprieta PDM in riga.
- **`core/properties_manager.py`** - mapping normalizzato esteso ai nuovi campi e sync PDM->SW aggiornata per esportare `code`, `state`, `created_by` (nome utente) e `created_at`.
- **`config.py`** - versione applicazione aggiornata a `2.10.16`.

### v2.10.15
**Mappature proprieta PDM <-> SolidWorks + sync centralizzata**

- **`ui/sw_config_dialog.py`** - aggiunta tab **`Mappature proprieta`** in Configurazione SolidWorks per mappare i campi PDM fondamentali (`revision`, `title`, `description`) con i nomi proprieta SolidWorks, direzione sync e flag `Forza PDM`.
- **`config.py`** - aggiunta chiave profilo `sw_property_mapping` (mappature diverse per ambiente/profilo).
- **`core/properties_manager.py`** - introdotti metodi centralizzati `get_property_mapping`, `resolve_property_owner_doc`, `sync_sw_to_pdm`, `sync_pdm_to_sw`.
- **`core/properties_manager.py`** - regola DRW: i campi fondamentali PDM usano come owner il PRT/ASM padre (quando disponibile); le custom del DRW restano importate in `Proprieta SW` del documento.
- **`ui/document_dialog.py`** - `_import_props_from_sw` ora usa la sync centralizzata invece del salvataggio raw.
- **`macros/sw_bridge.py`** - pre-checkin aggiornato: import SW->PDM + enforce revisione PDM->SW prima dell'archiviazione.

### v2.10.14
**Nuova tab dedicata ai non codificati**

- **`ui/main_window.py`** - aggiunta terza tab **`🗂️ Non codificati`** oltre a Archivio CAD e Workspace.
- **`ui/archive_view.py`** - refactor con modalita `view_mode`: la tab Archivio CAD mostra solo codificati, la nuova tab mostra solo non codificati.
- **`ui/main_window.py`** - toolbar (Checkout/Check-in/Workflow) abilitata anche quando è attiva la tab Non codificati.

### v2.10.13
**Archivio CAD: separazione record non codificati**

- **`ui/archive_view.py`** - `refresh`: i documenti non codificati non vengono piu mostrati nei nodi codice dell'Archivio CAD principale.
- **`ui/archive_view.py`** - aggiunta cartella top-level **"📁 Non codificati"** con elenco dedicato dei record non codificati.
- **`ui/archive_view.py`** - conteggio in footer aggiornato con metrica separata `Non codificati: N`.

### v2.10.12
**Visualizzazione BOM gerarchica**

- **`ui/document_dialog.py`** - `_refresh_bom`: la BOM viene mostrata in ordine gerarchico DFS con indentazione dei sottoassiemi/sottocomponenti.
- **`ui/document_dialog.py`** - `_del_component`: rimozione corretta della relazione selezionata anche quando la riga appartiene a un livello annidato (usa `parent_id, child_id` della riga).
- **`ui/detail_panel.py`** - tab Struttura aggiornata con vista gerarchica read-only coerente col dialog documento.

### v2.10.11
**Fix import ASM annullata su riferimenti non codificati**

- **`ui/asm_import_wizard.py`** - `_validate_asm_references`: la validazione hard ora controlla solo riferimenti che richiedono rinomina file (componenti codificati), evitando false anomalie sui non codificati.
- **`ui/asm_import_wizard.py`** - `_fix_asm_references`: il replace forzato dei riferimenti viene applicato solo ai file effettivamente rinominati.
- Risolto il caso mostrato in UI: import non più annullata quando restano riferimenti ai componenti non codificati.

### v2.10.10
**Import struttura ASM: BOM completa anche per non codificati**

- **`ui/asm_import_wizard.py`** - aggiunto `_ensure_bom_document_map`: prima del link BOM, ogni nodo dell'ASM viene mappato a un documento PDM (riuso se esiste, creazione automatica documento non codificato se assente).
- **`ui/asm_import_wizard.py`** - `_import`: la creazione relazioni BOM ora include tutto il contenuto dell'assieme, anche i componenti non selezionati per codifica.
- **`ui/asm_import_wizard.py`** - messaggio finale arricchito con conteggi: non codificati creati/riusati nel DB.

### v2.10.9
**Fix import token `SW-*` non espressi come `$PRP`**

- **`core/properties_manager.py`** - `_read_prop_value`: aggiunto tentativo primario con `Get6/Get5` byref per ottenere il valore valutato reale in modo generico (copre i token SW nella maggior parte dei casi).
- **`core/properties_manager.py`** - aggiunti `_parse_sw_token_expression` e `_resolve_sw_token_expression` per risolvere token `SW-*` grezzi (es. `"SW-Mass@..."`) quando restano non valutati.
- **`core/properties_manager.py`** - supporto esplicito a `SW-Mass`, `SW-Volume`, `SW-Density`, `SW-SurfaceArea` via `CreateMassProperty()` come fallback.

### v2.10.8
**Fix import proprieta linkata `SW-File Name`**

- **`core/properties_manager.py`** - `_read_summary_info`: aggiunta gestione campi file-based (`SW-File Name` e alias) risolti dal path/titolo del documento SolidWorks, evitando il salvataggio dell'espressione `$PRP:"SW-File Name"`.
- **`core/properties_manager.py`** - supportata anche variante senza estensione (`File Name without extension`) quando presente nei link.

### v2.10.7
**Fix esteso import proprieta linkate SolidWorks**

- **`core/properties_manager.py`** - parser link expression generalizzato: supporta `$PRP:"..."`, `$PRPSHEET:"..."`, varianti con apici singoli e target senza apici.
- **`core/properties_manager.py`** - risoluzione SummaryInfo estesa: supporta alias sia con prefisso `SW-` sia standard (`Author`, `Title`, `Subject`, `Company` tramite custom, ecc.).
- **`core/properties_manager.py`** - supporto target con suffisso `@config/file` e risoluzione cross-config tramite lookup su più `CustomPropertyManager`.

### v2.10.6
**Fix import proprieta linkate SolidWorks (`$PRP`)**

- **`core/properties_manager.py`** - aggiunta risoluzione esplicita delle espressioni `$PRP:"..."`/`$PRPSHEET:"..."` durante l'import, con supporto ai campi SummaryInfo (`SW-Author`, `SW-Title`, ecc.) e fallback su custom property target.
- **`core/properties_manager.py`** - `_read_custom_props`: se il valore letto e una link expression, ora viene risolto al valore finale prima del salvataggio nel DB.

### v2.10.5
**Fix import custom properties da SolidWorks (valore valutato)**

- **`core/properties_manager.py`** - `_read_custom_props`: la lettura proprieta ora privilegia il valore risolto (evaluated) invece dell'espressione raw (`$PRP:"..."`) quando disponibile.
- **`core/properties_manager.py`** - aggiunti helper `_read_prop_value`, `_best_value_from_result`, `_is_link_expr` con fallback compatibile tra API COM SW (`Get6`, `Get5`, `Get4`, `Get2`, `Get`).

### v2.10.4
**Import ASM: copia in workspace anche dei componenti non codificati**

- **`ui/asm_import_wizard.py`** - `_fallback_copy_files`: esteso il fallback (quando Pack and Go non e disponibile) per copiare in workspace anche i file SolidWorks non selezionati per la codifica, mantenendo il nome originale.
- **`ui/asm_import_wizard.py`** - `_fallback_copy_files`: aggiunti controlli anti-collisione con file codificati e deduplica per evitare sovrascritture/duplicati.

### v2.10.3
**Fix apertura documento in SolidWorks (no nuova sessione duplicata)**

- **`ui/detail_panel.py`** - `_open_in_solidworks`: sostituita l'apertura diretta via `subprocess` su `SLDWORKS.exe` con apertura COM (`GetActiveObject`/`Dispatch` + `OpenDoc`) per riusare la sessione SW già aperta quando disponibile.
- **`ui/detail_panel.py`** - `_open_in_solidworks`: mantenuto fallback a eseguibile configurato / `cmd start` solo se COM non disponibile.

### v2.10.2
**Wizard import ASM: coerenza selezione e preview codici**

- **`ui/asm_import_wizard.py`** - `_is_row_checked` / `_set_row_checked` / `_on_check_changed`: radice ASM resa obbligatoria in ogni scenario (anche con Seleziona tutti / Deseleziona tutti), con ripristino forzato dello stato.
- **`ui/asm_import_wizard.py`** - `_refresh_preview_codes` + `_update_row_code`: anteprima codici consolidata in modalita sequenziale simulata globale, evitando codici duplicati in tabella durante la selezione.
- **`ui/asm_import_wizard.py`** - `_apply_global`, `_select_all`, `_deselect_all`: refresh preview unico a fine operazione e mantenimento evidenziazione righe attive con checkbox selezionata.

---
### v2.10.1
**Importazione ASM fail-safe (no codici inutilizzabili)**

- **`ui/asm_import_wizard.py`** - `_import`: introdotto flusso fail-safe con rollback completo in caso di errore (file workspace/archivio, documenti DB creati nella sessione e contatori codifica).
- **`ui/asm_import_wizard.py`** - `_import`: validazione hard dei riferimenti interni ASM prima del commit finale; se resta anche un solo riferimento ai file originali l'import viene annullato.
- **`ui/asm_import_wizard.py`** - `_import`: archiviazione limitata ai soli documenti creati nell'import corrente (i documenti gia presenti nel PDM non vengono sovrascritti).
- **`ui/asm_import_wizard.py`** - `_run_pack_and_go`: aggiunto fallback di chiamata COM `GetPackAndGo(status)` per compatibilita con versioni SW in cui il parametro stato e richiesto.
- **`ui/asm_import_wizard.py`** - `_fix_asm_references`: sostituzione riferimenti tramite `ISldWorks.ReplaceReferencedDocument` con documento ASM chiuso (modalita piu affidabile per COM).

---
### v2.10.0
**Pannello dettaglio nodo-codice rinnovato + fix thumbnail**

- **`ui/detail_panel.py`** – `_build_code_panel` completamente ridisegnato:
  - Card **Parte/Assieme**: thumbnail 110×110 con fallback icona tipo, label info (Rev., stato), pulsante **📤 Esporta in WS** (visibile solo se file in archivio e NON già in workspace), pulsante **🔨 Apri in SolidWorks** (visibile se file disponibile in archivio o workspace).
  - Card **Disegno**: stessa struttura + pulsante **📐 Aggiungi DRW** (visibile solo se PRT/ASM archiviato e DRW assente).
  - Sezione **Azioni** (Crea in SW / Crea da file) rimane ma è visibile solo quando il PRT/ASM non ha ancora un file in archivio.
  - Nuovo helper `_load_thumb_to(doc, lbl)`: carica thumbnail in un label target generico (riutilizzabile per ambedue le card).
  - Nuovo helper `_export_doc_to_ws(doc)`: copia file da archivio a workspace con `shutil.copy2`, poi ricarica il pannello per aggiornare i pulsanti.
  - Nuovo helper `_open_in_solidworks(doc)`: cerca il file prima in workspace poi in archivio, apre con `SLDWORKS.exe` configurato (fallback a `cmd start`).
  - Nuovo helper `_is_in_workspace(doc)` e `_get_archive_file(doc)`.
  - Fix `_load_thumbnail`: rimossa doppia chiamata ridondante `setText`.

---

### v2.9.9
**Fix wizard importazione ASM**

- **`core/asm_manager.py`** – `_read_components_recursive`: rimosso il fallback `except TypeError: raw = model.GetComponents` che restituiva l'oggetto metodo anziché il risultato (causa principale del blocco al primo livello). Aggiunto `sw.GetOpenDocumentByName(fp)` come step intermedio prima di `OpenDoc6`. Log warning invece di `pass` silenzioso per facilitare il debug.
- **`ui/asm_import_wizard.py`** – `_set_row`: la colonn **Livello** ora usa il default per tipo (depth=0 → LIV0-Macchina forzato e non modificabile, ASM → LIV1-Gruppo, PRT → LIV2-Parte) indipendentemente dalla selezione globale. La radice ASM ha checkbox disabilitata (sempre inclusa) e combo livello disabilitata.
- **`ui/asm_import_wizard.py`** – `_fill_group_combo`: aggiunto `blockSignals(True/False)` per evitare che il `clear()` generi eventi `currentIndexChanged` fantasma durante la costruzione delle righe.
- **`ui/asm_import_wizard.py`** – `_populate_table`: aggiunto refresh forzato di tutti i codici proposti dopo il completamento della tabella, garantendo che tutte le righe (non solo la prima) mostrino il codice proposto.

---

### v2.9.5
**Fix Pack&Go multi-livello**

- **`core/asm_manager.py`** – `_read_components_recursive`: quando `GetModelDoc2()` restituisce `None` (componente **lightweight**), apre il sotto-assieme silenziosamente in SW (`OpenDoc6` silent+read-only), ne legge la struttura e poi lo chiude. La ricorsione ora raggiunge tutti i livelli indipendentemente dalla modalità di caricamento dei componenti. Parametro `sw` (istanza SolidWorks) propagato lungo tutta la catena ricorsiva.
- **`ui/asm_import_wizard.py`** – `_run_pack_and_go`: parsing robusto della tupla COM restituita da `GetDocumentNames()` (gestisce sia 2-tuple che 3-tuple in base a versione SW/win32com). Controllo esplicito sul risultato di `Save()` con errore chiaro se nessun file viene creato.

---

### v2.9.4
**Wizard importazione massiva struttura ASM da SolidWorks**

- **`core/asm_manager.py`**: aggiunto `read_asm_tree_from_active()` — legge ricorsivamente la struttura di un assieme aperto in SolidWorks via COM, senza scrivere nulla nel DB. Ritorna lista piatta DFS con `{name, path, type, depth, parent_path, quantity}`.
- **`ui/asm_import_wizard.py`** (NUOVO): dialog wizard `AsmImportWizard`. Mostra l'albero dell'assieme in una tabella con per ogni componente: checkbox inclusione, combo macchina/gruppo/livello per row, anteprima codice auto-aggiornata. Impostazioni globali con "Applica a tutti". Riconosce automaticamente i documenti già presenti nel PDM. Alla conferma: genera i codici, crea i documenti nel DB, esegue **Pack and Go** via COM (copia workspace con file rinominati e riferimenti aggiornati), archivia i file nel PDM e collega la BOM.
- **`macros/pdm_asm_import.py`** (NUOVO): entry point Qt per il wizard, lanciabile dalla macro VBA in modalità non-bloccante.
- **`macros/PDM_Integration.swb`**: aggiunto sub `PDM_ImportAsm()` e helper `RunAsmImport()` — rileva il file `.SLDASM` attivo e lancia il wizard.
- **`ui/main_window.py`**: aggiunto menu **SolidWorks → 📦 Importa struttura ASM da SW…** come alternativa alla macro VBA.

---

### v2.9.3
**Fix sicurezza autenticazione, backup DB automatico, export BOM Excel**

- **`core/user_manager.py`**: fix vulnerabilità autenticazione — utenti con `password_hash` vuoto non possono più accedere con password arbitrarie; ora è richiesta la corrispondenza esatta dell'hash.
- **`core/backup_manager.py`** (NUOVO): gestore backup database. Crea copie timestamped di `pdm.db` in `<shared_root>/database/backups/`, mantiene gli ultimi 10 backup con rotazione automatica. Supporta anche `restore()` da backup con salvataggio preventivo del DB corrente.
- **`ui/main_window.py`**: aggiunta voce menu **Strumenti → 💾 Backup database…** che esegue un backup immediato e mostra il percorso del file creato.
- **`ui/document_dialog.py`**: aggiunto pulsante **📊 Esporta BOM Excel** nel tab BOM — esporta la BOM appiattita (tutti i livelli) in un file `.xlsx` con intestazione colorata e larghezze colonne ottimizzate.

---

### v2.9.2
**Archive view redesign & creazione archive-first**

- **`ui/archive_view.py`**: i nodi codice diventano **selezionabili** (prima erano non-selezionabili).
  I figli (PRT/ASM/DRW) sono mostrati **solo** se `archive_path` è valorizzato **o** il doc è in checkout (`is_locked=1`).
  Il filtro per tipo opera solo sui figli — i nodi codice rimangono sempre visibili.
- **`ui/detail_panel.py`**: aggiunta **doppia modalità**:
  - *Modalità documento*: comportamento precedente (thumbnail, tab info/props/BOM/storico).
  - *Modalità codice*: mostra i tipi disponibili (PRT/ASM/DRW) con relativo stato, e pulsanti azione:
    - `[Crea in SW]` — crea dal template SW ed archivia direttamente (popup opzionale checkout)
    - `[Crea da file]` — importa un file esterno nell'archivio
    - `[Aggiungi DRW]` — crea disegno DRW nell'archivio (visibile solo se PRT/ASM presente)
  - Checkbox **"Metti in checkout dopo la creazione"** permanente nel gruppo azioni.
- **`core/file_manager.py`**: nuovo metodo `create_to_archive(document_id, source_path=None)`.
  Crea il file direttamente in archivio (flusso *archive-first*) — opzionalmente da template SW o
  da file esterno. Aggiorna `archive_path` nel DB. Il checkout rimane opzionale e separato.
- **`core/database.py`**: migrazione automatica colonna `pdf_path TEXT` su `documents`
  (preparazione generazione PDF al checkin).

---

### v2.9.1
**Sola lettura fisica su file workspace**

- **`core/checkout_manager.py`**: aggiunto `_set_readonly()` e `_set_writable()` (helper statici con `stat.S_IWRITE`).
- `_copy_archive_to_workspace()`: ogni copia da archivio → workspace è ora **sola lettura** per default.
  Questo vale per: consultazione, componenti ASM, copia mentre documento in checkout altrui.
- `checkout()`: dopo la copia rende il file **scrivibile** — solo il proprietario del lock può modificarlo.
- `checkin()` / `undo_checkout()` / `remove_from_workspace()`: `_set_writable()` chiamato prima di `unlink()`
  (su Windows un file readonly non può essere eliminato).

**Regola di sola lettura nella workspace:**

| Situazione | File in WS | Scrivibile |
|---|---|---|
| Checkout mio | ✅ | ✅ |
| Consultazione | ✅ | ❌ |
| Componente ASM (copia) | ✅ | ❌ |
| Checkout altrui (copia archivio) | ✅ | ❌ |
| Rilasciato / Obsoleto (copia archivio) | ✅ | ❌ |

---

### v2.9.0
**Distribuzione portabile – primo avvio con cartella dati separata**

- **`config.py`**: `local_config.json` non è più nella cartella sorgenti — viene cercato nella
  cartella dati locali indicata in `.pdm_datadir`. Fallback a `APP_DIR` per compatibilità.
  Migrazione automatica `_init_workspace` → profilo al primo salvataggio.
- **`ui/first_run_dialog.py`** (NUOVO): dialog al primo avvio che chiede (1) cartella dati locali
  e (2) workspace SolidWorks. Scrive `.pdm_datadir` e inizializza `local_config.json`.
- **`main.py`**: check presenza `.pdm_datadir` all'avvio; se assente mostra `FirstRunDialog`.
- **`make_dist.py`** (NUOVO): script che crea `dist/PDM-SW-2/` escludendo `.venv/`,
  `local_config.json`, `.pdm_datadir` e file runtime. Opzione `--zip` per archivio compresso.
- **`install.bat`** / **`start.bat`** (NUOVI): script per installazione e avvio su nuovo PC.
- **`.gitignore`** (NUOVO): esclude file locali e di runtime dal controllo versione.

---

### v2.8.1
**Bugfix: dialog "Genera Codice" non si apriva premendo il pulsante**

- **`macros/pdm_panel.py`** – `CreateCodeDialog._build_ui()`: rimossa chiamata prematura a `_update_level_options()` durante la costruzione di `cmb_level`, che causava un `AttributeError` silenzioso su `lbl_preview` (non ancora creato) — con `pythonw.exe` l'eccezione era invisibile e il dialog non si apriva mai

---

### v2.8.0
**Genera Codice → Save As + checkout automatico nella workspace**

- **`macros/pdm_panel.py`** – `CreateCodeDialog._create()` riscritta:
  - Dopo aver generato il codice PDM, esegue **Save As** del documento aperto in SolidWorks via COM (`SaveAs3`) direttamente nella workspace locale con nome `CODICE.EXT`
  - Fallback automatico a `shutil.copy2` se SolidWorks COM non è disponibile
  - Documento bloccato immediatamente come **checkout** (update `is_locked`, insert `checkout_log`, insert `workspace_files`)
  - Il pannello aggiorna `file_path` al nuovo percorso e si rinfresca
- **`_do_generate_code()`**: aggiornato per propagare il nuovo `file_path` dal dialogo al pannello

---

### v2.7.0
**Pannello macro Qt – Genera Codice + UI pulsanti**

- **`macros/pdm_panel.py`** (NUOVO): pannello PyQt6 lanciato dalla macro VBA, con pulsanti colorati per ogni azione PDM
  - Header con badge stato, lock, tipo documento e revisione
  - Pulsanti: Checkout, Check-in, Annulla CO, Consultazione, Proprietà↑, Proprietà↓, Apri PDM
  - Sezione **Genera Codice PDM** (visibile solo se file non nel PDM): form con macchina/gruppo/livello/titolo e preview codice in tempo reale
  - Salvataggio automatico del file SW via COM prima del Check-in
  - Rimane "sempre in primo piano" (WindowStaysOnTopHint)
- **`sw_bridge.py`**: aggiunte azioni `panel` (lancia pdm_panel.py in background) e `create_code` (genera codice da JSON parametri)
- **`PDM_Integration.swb`** v3.0 — rimpiazza l'InputBox con numeri:
  - `main()` / `PDM_Panel()` aprono il pannello Qt non-bloccante
  - `PDM_Checkout`, `PDM_Checkin`, `PDM_UndoCheckout`, `PDM_OpenApp` rimangono come shortcut toolbar
  - Nuova `RunPanel` (launch asincrono, `bWait=False`) vs `RunBridge` (sincrono)
  - `GetActiveFilePathSilent()` per open panel senza MsgBox di errore

### v2.6.0
**Integrità file – Overwrite warning e verifica fisica R7**

- **Checkout – avviso sovrascrittura**: prima del checkout, confronto MD5 tra file in workspace e archivio; se diversi, chiede conferma sovrascrittura (`No` di default)
- **Esporta in WS – avviso sovrascrittura**: stessa logica MD5 nel dialog "Esporta in WS" della scheda documento
- **R7b – verifica fisica**: `change_state(shared_paths=sp)` verifica l'esistenza fisica del file su disco prima del rilascio (non solo il campo `archive_path` nel DB)
- `_propagate_state_to_companion()` e `change_state()` accettano ora `shared_paths` opzionale per propagare il controllo fisico al companion DRW/PRT
- `workflow_dialog.py` passa `session.sp` a `change_state()`

### v2.5.0
**Workflow v3 – Semplificazione a 4 stati**

- **Eliminato stato `Revisionato`**: workflow semplificato a 4 stati: In Lavorazione, Rilasciato, In Revisione, Obsoleto
- **Nuovo flusso**: In Lavorazione → Rilasciato (00) → Crea revisione → In Revisione → Rilasciato (01) → ...
- **Crea revisione**: operazione (non transizione) da stato Rilasciato — crea nuovo documento in stato In Revisione
- **Annulla revisione**: nuovo comando nel menu contestuale — elimina revisione In Revisione e torna alla precedente
- **Guard ultima revisione**: cambio stato consentito solo sull'ultima revisione di un codice
- **Workflow dialog semplificato**: rimosso bottone "⚡ Rilascia documento", mostra solo transizioni consentite
- **Menu contestuale riordinato**: Consultazione, Checkout, Checkin, Annulla checkout, Crea revisione, Annulla revisione, Workflow, Apri in eDrawings, Proprietà
- **READONLY_STATES** aggiornato: solo `Rilasciato` e `Obsoleto` (In Revisione è modificabile)
- Rimosso `release_document()` da workflow_manager (ridondante con `change_state`)

### v2.4.0
**PDM Profile – Gestione multi-ambiente**

- **Profili multi-ambiente**: supporto per N profili di lavoro con database, archivio e configurazione SW indipendenti (es. clienti diversi, versioni software diverse)
- **Selettore profilo all'avvio**: se esiste più di un profilo, dialog di selezione all'avvio con combo + pulsante "Gestisci profili"
- **Dialog gestione profili** (`PDM Profile`): crea, rinomina, elimina, copia profili con visualizzazione dettagli (exe, template, workspace)
- **Copia profilo**: due modalità — "Solo configurazione" (copia impostazioni SW) o "Configurazione + dati" (copia anche archivio/thumbnail, escluso database)
- **Cambio profilo a runtime**: menu Strumenti → PDM Profile, riconnessione al DB del nuovo profilo con ri-autenticazione automatica dell'utente
- **Profilo nella status bar**: indicatore del profilo attivo visibile in basso a sinistra
- **Migrazione automatica**: config flat esistente migrata automaticamente in un profilo "Default" al primo avvio
- **Backward-compatible**: `load_local_config()` / `save_local_config()` restano identiche nell'API, ora profile-aware internamente
- Nuovo file `ui/profile_dialog.py`

### v2.3.0
**Configurazione eseguibili SolidWorks/eDrawings + gestione registro selettiva**

- **Percorsi eseguibili configurabili**: campi dedicati per SolidWorks e eDrawings nella configurazione, con pulsanti "Sfoglia" e "Rileva" (auto-detect da registro Windows)
- **Apri in eDrawings migliorato**: usa l'eseguibile configurato invece di cercare percorsi hardcoded; se non configurato, mostra avviso con istruzioni
- **Apri in SolidWorks migliorato**: apertura file tramite eseguibile configurato con `subprocess.Popen` anziché `os.startfile`
- **Gestione registro selettiva**: import `.reg`/`.sldreg` con scelta categorie (opzioni sistema, toolbar, scorciatoie tastiera, gesture mouse, personalizzazioni menu, viste salvate)
- **Parser registro**: nuovo modulo `core/reg_manager.py` con parsing blocchi, categorizzazione chiavi e scrittura filtrata
- Nuovo file `core/reg_manager.py`

### v2.2.0
**Pannello dettaglio documento + integrazione eDrawings**

- **DetailPanel**: pannello laterale nell'archivio CAD con anteprima readonly del documento selezionato (info, proprietà SW, BOM, storico)
- **QSplitter ridimensionabile**: albero documenti e pannello dettaglio separati da splitter trascinabile (65/35 default, collapsabile)
- **Thumbnail**: visualizzazione anteprima immagine da `{shared_root}/thumbnails/` con fallback a icona tipo documento
- **Generazione thumbnail al checkin**: export automatico PNG via eDrawings dopo ogni check-in (non bloccante, fallback silenzioso)
- **Apri in eDrawings**: bottone nel pannello dettaglio + voce nel menu contestuale per consultazione rapida senza aprire SolidWorks
- **Selezione → anteprima**: singolo clic aggiorna il pannello dettaglio; doppio clic continua ad aprire il dialog completo
- Nuovo file `ui/detail_panel.py`

### v2.1.2
**Fix verifica errori**

- **Bug fix**: ripristinato metodo `FileManager.export_to_workspace()` — era stato rimosso ma ancora referenziato in `DocumentDialog._esporta_ws()`
- **Type hints**: corretti parametri `Optional[int]` e `Optional[set]` in `DocumentDialog` e `AsmManager`

### v2.1.1
**Elenco componenti nei dialog checkout/checkin per assiemi**

- **Checkout dialog ASM**: aggiunta tabella con elenco componenti (codice, tipo, revisione, titolo) al posto del semplice contatore
- **Checkin dialog ASM**: aggiunta tabella BOM struttura assieme (con quantità) sopra la tabella dei file in workspace
- Migliorata leggibilità: l'utente vede esattamente quali file sono coinvolti nell'operazione

### v2.1.0
**Sync automatico proprietà e BOM al check-in + fix import manuale**

- **Check-in automatico**: ogni check-in da SolidWorks ora importa automaticamente le proprietà custom dal file aperto nel PDM
- **Check-in ASM**: per i documenti Assieme, il check-in aggiorna anche la struttura BOM dal file attivo
- **Blocco check-in**: se l'importazione proprietà/BOM fallisce (file non aperto in SW), il check-in viene bloccato con messaggio esplicito
- **BOM import manuale**: il bottone "Importa struttura da SW" non chiede più il file — legge direttamente dal documento attivo in SolidWorks
- **Nuovo metodo**: `AsmManager.import_bom_from_active_doc()` — legge BOM via `GetActiveObject` + `ActiveDoc.GetComponents`
- **Fix proprietà COM**: `read_from_sw_file()` ora usa `ActiveDoc` come primo tentativo, `GetNames` come proprietà (non metodo), `Get()` al posto di `Get5()` (che fallisce con GetActiveObject)
- **Fix errori silenziosi**: errore `_error` ora viene mostrato all'utente con `QMessageBox.warning` anziché essere filtrato silenziosamente
- **Fix salvataggio**: le proprietà importate manualmente vengono ora salvate nel DB e la tabella ricaricata

### v2.0.7 (merged)
- Fix bottone "Importa da SW" nel dialog Proprietà – non faceva nulla

### v2.0.6
**Fix trasferimento proprietà SolidWorks ↔ PDM**

- **Root cause import**: Python tentava `OpenDoc6` su un file già aperto in SolidWorks → falliva silenziosamente → 0 proprietà lette
- **Root cause export**: `OpenDoc6` con parametri `int` per ByRef → errore COM "Incompatibilità tra tipi"
- **Soluzione import (azione 4)**: La VBA legge le proprietà direttamente dal modello aperto, le scrive in `sw_props_temp.json`, Python le salva nel DB. Nuovo pattern VBA→JSON→Python che bypassa completamente i problemi COM cross-process
- **Soluzione export/read COM**: `GetActiveObject` + `GetOpenDocumentByName` per riusare il doc già aperto, `OpenDoc6` con `VARIANT(VT_BYREF|VT_I4)` per parametri ByRef corretti, `Get5` con fallback a `Get` per versioni SW vecchie
- Nuova azione `import_props_json` nel bridge Python
- Errori `_error` non più silenziati (venivano rimossi con `pop`)
- Logging dettagliato su ogni step di import/export
**Fix comunicazione macro: MessageBox non visibile da SolidWorks**

- **Root cause**: `MessageBoxW` chiamata dal processo Python lanciato con window mode `0` (nascosto) non viene visualizzata in SolidWorks. Il processo gira, l'azione viene eseguita, ma l'utente non vede alcun feedback.
- **Soluzione**: Nuovo pattern di comunicazione result-file:
  - `sw_bridge.py` scrive il risultato in `macros/sw_bridge_result.txt` (prima riga `OK`/`ERR`, resto = messaggio)
  - `PDM_Integration.swb` legge il file dopo che il processo termina e mostra `MsgBox` nel contesto di SolidWorks (dove è **garantito** che sia visibile)
- `sw_bridge.py`: mantenuto logging su file (`macros/sw_bridge.log`) per debug
- `sw_bridge.py`: `MessageBoxW` con `MB_TOPMOST` come fallback (se non in SolidWorks)
- `PDM_Integration.swb`: cancella file risultato prima di ogni esecuzione
- `PDM_Integration.swb`: lettura file Unicode con `TristateTrue`

### v2.0.3
**Fix macro SolidWorks: silent failure da SolidWorks**

- `sw_bridge.py`: aggiunto **logging su file** (`macros/sw_bridge.log`) — cattura ogni errore anche con `pythonw.exe` (che non ha console)
- `sw_bridge.py`: wrapper `main()` con try/except globale che scrive traceback completo nel log
- `PDM_Integration.swb`: `wsShell.CurrentDirectory = appDir` — **setta CWD** alla directory PDM prima di eseguire (SolidWorks imposta CWD alla propria dir di installazione)
- `PDM_Integration.swb`: usa `python.exe` (non `pythonw.exe`) come priorità per debug visibile
- `PDM_Integration.swb`: rimosso `fso.GetAbsolutePathName(".")` (inutilizzabile dentro SolidWorks)
- `PDM_Integration.swb`: quoting robusto con `Chr(34)` per percorsi con spazi
- Aggiunto `macros/test_macro.bat` per test manuale del bridge
- Messaggio errore VBA ora include path del file log

### v2.0.2
**Fix macro SolidWorks (sw_bridge.py + PDM_Integration.swb)**

- `sw_bridge.py` aggiornato all'API v2.0: `checkout()` e `checkin()` usano la nuova firma senza parametri path
- Aggiunta azione `undo_checkout` (mancante)
- Lookup documento per codice+tipo (dal nome file workspace) invece che solo `file_name`
- Usa `pythonw.exe` dal venv (`.venv\Scripts\pythonw.exe`) invece del sistema
- `PDM_Integration.swb`: corretto passaggio parametri ai sub, aggiunto `PDM_UndoCheckout`
- Menu macro aggiornato con tutte le 6 azioni disponibili
- Verifica esistenza `sw_bridge.py` prima di lanciare il comando

### v2.0.1
**Bugfix: "no such table: main._documents_old" durante checkin/checkout**

- **Root cause**: SQLite >= 3.25 aggiorna automaticamente i FK in tutte le tabelle durante `ALTER TABLE RENAME`. La migration del UNIQUE constraint rinominava `documents` in `_documents_old`, corrompendo i FK di `checkout_log`, `workflow_history`, `asm_components` ecc.
- **Fix 1 – Riparazione DB**: nuovo metodo `_repair_stale_fk_references()` che usa `PRAGMA writable_schema` per correggere direttamente i riferimenti corrotti in `sqlite_master` all'avvio
- **Fix 2 – Prevenzione**: aggiunto `PRAGMA legacy_alter_table=ON` prima del RENAME per impedire la propagazione automatica dei FK
- **Fix 3 – Schema completo**: `NEW_TABLE_DDL` nella migration ora include le colonne `checkout_md5/size/mtime`; la copia dati gestisce dinamicamente le colonne presenti nella vecchia tabella

### v2.0.0
**Revisione completa checkout/checkin e gestione archivio/workspace**

#### Checkout / Checkin
- **Blocco checkout/checkin** su stati `Rilasciato`, `Obsoleto` (costante `READONLY_STATES`)
- **Snapshot MD5** al checkout: al checkin confronta hash per rilevare modifiche reali e conflitti
- **Checkout ASM ricorsivo**: copia tutti i componenti BOM in workspace senza lock (role `component`)
- **Checkbox DRW**: opzione per includere il Disegno associato nel checkout
- **Consultazione**: copia in workspace senza lock, registrata come role `consultation`
- **Annulla checkout**: solo l'admin puo revocare il checkout di altri utenti
- **Tabella workspace_files**: tracking persistente di tutti i file in workspace per utente

#### Workflow
- **Propagazione bidirezionale PRT/ASM <-> DRW**: cambio stato sincronizzato automaticamente (stessa revisione)
- **Auto-obsolescenza**: rilascio di nuova revisione rende obsolete le revisioni precedenti rilasciate
- **Nuova revisione**: copia file archiviato come base + crea automaticamente revisione DRW companion
- **cancel_revision()**: eliminazione revisione non rilasciata con pulizia file archivio

#### File Manager
- **import_file()**: accetta file solo dalla workspace configurata (no percorsi esterni)
- **create_from_external_file()**: unico entry-point per file esterni, copia in workspace con codice PDM
- **export_from_workspace()**: esporta dalla workspace (non dall'archivio direttamente)
- Rimossi `export_file()` e `export_to_workspace()` (sostituiti)

#### UI
- **Archivio CAD**: QTreeWidget raggruppato per codice, nodi padre = codice, figli = tipo+revisione
- **Filtro tipo**: agisce solo sui nodi figlio (nasconde tipi non corrispondenti)
- **Workspace view**: due sezioni separate — "In Checkout (modificabili)" e "In Workspace (copie/consultazione)"
- **CheckoutDialog**: dialog conferma checkout con opzione DRW, info componenti ASM
- **CheckinDialog**: dialog checkin con rilevamento modifica (verde/giallo/rosso), tabella ASM selezionabile
- **Menu contestuale archivio**: Checkout, Checkin, Annulla checkout, Consultazione, Esporta da workspace, Cambia stato, Nuova revisione, Proprieta

#### Database
- Nuove colonne `documents`: `checkout_md5`, `checkout_size`, `checkout_mtime`
- Nuove colonne `checkout_log`: `checkout_md5`, `checkout_size`, `checkout_mtime`
- Nuova tabella `workspace_files`: tracking file in workspace per utente con ruolo e parent

### v1.4.1
**Bugfix: recovery da migration DB rotta**

- `_migrate_documents_unique()` riscritta con gestione dei tre stati possibili:
  - **Recovery**: se `_documents_old` esiste ma `documents` no (crash a metà migration precedente) → completa automaticamente senza RENAME
  - **Pulizia**: se entrambe esistono (migration già completata) → drop del residuo
  - **Normale**: migration completa con rollback esplicito in caso di errore (rimosso `pass` silenzioso)
- Estratto helper `_create_doc_indexes()` per riuso

### v1.4.0
**Architettura workspace-first – i file SI aprono e si salvano SOLO nella workspace locale:**

- **Crea in SW**: crea file da template nella workspace; per PRT/ASM propone opzione DRW
- **Crea da file** (NUOVO): seleziona file da qualsiasi cartella, lo copia in workspace con codice PDM; per PRT/ASM cerca DRW nella stessa cartella
- **Importa in PDM** (ex "Importa File"): cerca `{codice}.SLDPRT/ASM/DRW` nella workspace e lo archivia; per PRT/ASM importa automaticamente anche il DRW se presente
- **Esporta in WS**: copia dall’archivio PDM nella workspace (non più cartella libera)
- **Apri in SW**: copia dall’archivio in workspace (se non già presente) e apre — **mai apre dall’archivio direttamente**
- Estensione attesa determinata dal `doc_type` del documento (`Parte→.SLDPRT`, `Assieme→.SLDASM`, `Disegno→.SLDDRW`)
- Helper FileManager: `copy_to_workspace()`, `import_from_workspace()`, `open_from_workspace()`, `export_to_workspace()`

### v1.3.0
- **Template SolidWorks**: estensioni corrette `.prtdot`, `.asmdot`, `.drwdot` (con mapping → `.SLDPRT`, `.SLDASM`, `.SLDDRW`)
- **Crea in SW**: per PRT/ASM propone opzione "Crea anche DRW" → copia `.drwdot` e crea documento Disegno in DB con stesso codice
- **Importa file**: filtra automaticamente le estensioni in base al tipo documento; per PRT/ASM rileva automaticamente il DRW companion (stesso nome base) e propone di importarlo
- **Esporta file**: per PRT/ASM propone di esportare anche il DRW associato
- `is_code_available()`: permette stesso codice con `doc_type` diverso (Disegno può condividere codice col PRT/ASM padre)
- Nuovi metodi `FileManager`: `get_drw_document()`, `get_or_create_drw_document()`, `find_companion_drw()`

### v1.2.0
- Nuovo dialog **Configurazione SolidWorks** (Strumenti → Configurazione SolidWorks)
  - Template PRT / ASM / DRW configurabili per workstation
  - File `.reg` SolidWorks: scelta e applicazione al registro di Windows
  - Workspace locale: cartella di lavoro predefinita
- Pulsante **Crea in SW** nel dettaglio documento: copia il template, rinomina col codice PDM, apre in SolidWorks
- **Import file**: rinomina automaticamente con il codice PDM (`ABC_COMP-0001.SLDPRT`)
- **DRW (Disegno)**: livello dedicato nella creazione documento con link opzionale al documento padre (PRT/ASM)
- DB: aggiunto campo `parent_doc_id` (DRW → PRT/ASM)

### v1.1.1
- Fix crash tab Contatori (`session.current_user` → `session.user`)
- Tab Contatori carica i dati automaticamente all'apertura
- Descrizioni macchine e gruppi forzate in maiuscolo
- Revisioni formato numerico `00`, `01`, `02`…

### v1.1.0
- **Codifica gerarchica**: Macchine → Gruppi → Livelli (LIV0/LIV1/LIV2)
- Schema codici: `MMM_V001` (macchina), `MMM_GGGG-V001` (gruppo), `MMM_GGGG-0001` (parte PRT), `MMM_GGGG-9999` (sottogruppo ASM)
- Contatori versione indipendenti per macchina e per coppia macchina+gruppo
- Warning automatico collisione contatori LIV2 (parti↑ / sottogruppi↓)
- Nuovo dialogo "Configurazione Codifica" con tab Macchine / Gruppi / Contatori
- Flusso creazione documento aggiornato: Macchina → Livello → Gruppo → codice auto-generato
- Migrazione automatica DB (aggiunge `machine_id`, `group_id`, `doc_level` a documenti esistenti)

### v1.0.0
- Versione iniziale con archivio, checkout/checkin, workflow, macro VBA

---

## Caratteristiche principali

| Funzione | Dettaglio |
|---|---|
| **Architettura** | SQLite su cartella condivisa di rete – nessun server |
| **Utenti simultanei** | Max 5 (locking ottimistico su file) |
| **File supportati** | `.SLDPRT`, `.SLDASM`, `.SLDDRW` |
| **Workflow** | In Lavorazione → Rilasciato → In Revisione → Rilasciato (nuova rev) → Obsoleto |
| **Check-in / Checkout** | Con lock per file e storico completo |
| **Codifica gerarchica** | LIV0 Macchina → LIV1 Gruppo → LIV2 Parte/Sottogruppo |
| **Proprietà SW** | Import/Export da/verso SolidWorks via COM API |
| **BOM** | Gestione struttura assieme con import da file .SLDASM |
| **Macro VBA** | Integrazione diretta da SolidWorks |

---

## Struttura progetto

```
PDM-SW-2/
├── main.py                   # Entry point applicazione
├── setup_app.py              # Wizard configurazione iniziale
├── config.py                 # Configurazione globale
├── requirements.txt
├── core/
│   ├── database.py           # Gestione SQLite condiviso
│   ├── user_manager.py       # Utenti e autenticazione
│   ├── coding_manager.py     # Codifica documenti
│   ├── file_manager.py       # Gestione file archivio
│   ├── checkout_manager.py   # Check-in / Check-out
│   ├── workflow_manager.py   # Workflow stati
│   ├── properties_manager.py # Proprietà SolidWorks
│   └── asm_manager.py        # Struttura assieme (BOM)
├── ui/
│   ├── main_window.py        # Finestra principale
│   ├── session.py            # Sessione globale
│   ├── styles.py             # Tema dark
│   ├── setup_dialog.py       # Config percorso rete
│   ├── login_dialog.py       # Login utente
│   ├── archive_view.py       # Vista archivio CAD
│   ├── workspace_view.py     # Vista workspace utente
│   ├── document_dialog.py    # Dettaglio documento
│   ├── document_selector.py  # Selezione documento
│   ├── detail_panel.py       # Pannello dettaglio laterale
│   ├── profile_dialog.py     # PDM Profile: gestione profili
│   ├── coding_dialog.py      # Config codifica
│   ├── workflow_dialog.py    # Cambio stato workflow
│   └── users_dialog.py       # Gestione utenti
└── macros/
    ├── PDM_Integration.swb   # Macro SolidWorks (VBA/Basic)
    └── sw_bridge.py          # Bridge Python ↔ macro SW
```

---

## Installazione

### 1. Requisiti Python

```powershell
pip install -r requirements.txt
```

Pacchetti installati:
- `PyQt6` – interfaccia grafica
- `pywin32` – integrazione COM SolidWorks
- `Pillow` – anteprime immagini
- `openpyxl` – import/export Excel
- `filelock` – locking database concorrente

### 2. Prima configurazione (una sola volta)

Eseguire il wizard di setup:

```powershell
python setup_app.py
```

Il wizard:
1. Configura il percorso della cartella condivisa di rete
2. Crea la struttura delle cartelle PDM
3. Inizializza il database
4. Configura l'utente amministratore
5. Imposta la codifica documenti

### 3. Configurazione su PC aggiuntivi

Su ogni PC che usa il PDM, eseguire solo `main.py`:

```powershell
python main.py
```

Al primo avvio, viene chiesto il percorso della cartella condivisa.

---

## Avvio applicazione

```powershell
python main.py
```

---

## Struttura cartella condivisa

La cartella di rete condivisa avrà questa struttura:

```
\PDM\
├── database\
│   ├── pdm.db          # Database SQLite
│   └── pdm.lock        # File lock per scritture concorrenti
├── archive\
│   └── {CODICE}\
│       └── {REVISIONE}\
│           └── file.SLDPRT   # File archiviati
├── workspace\
│   └── {username}\           # Workspace per utente
├── thumbnails\
├── config\
└── temp\
```

---

## Workflow documenti

```
[Nuovo] → In Lavorazione → Rilasciato (rev.00)
                                │
                          Crea revisione
                                │
                          In Revisione → Rilasciato (rev.01)
                                              │
                                        Crea revisione → ...
                                        
Rilasciato (vecchia rev) → Obsoleto (automatico)
```

### Transizioni consentite:
| Da | A |
|---|---|
| In Lavorazione | Rilasciato |
| In Revisione | Rilasciato |
| Rilasciato | — (usare «Crea revisione» per nuova rev) |
| Obsoleto | — |

### Operazioni:
| Operazione | Da stato | Effetto |
|---|---|---|
| Crea revisione | Rilasciato | Crea nuova rev in «In Revisione» |
| Annulla revisione | In Revisione | Elimina la revisione, torna alla precedente |

---

## Integrazione SolidWorks

### Macro VBA (`macros/PDM_Integration.swb`)

1. In SolidWorks: **Strumenti → Macro → Esegui**
2. Selezionare `PDM_Integration.swb`

Oppure per accesso rapido:
1. **Strumenti → Personalizza → Barra strumenti**
2. Aggiungere pulsante macro

**Sub disponibili nella macro:**
- `PDM_Menu()` – menu interattivo con tutte le azioni
- `PDM_Checkout()` – checkout del file aperto
- `PDM_Checkin()` – salva e check-in del file aperto
- `PDM_ImportProperties()` – importa proprietà da SW nel PDM
- `PDM_ExportProperties()` – esporta proprietà PDM in SW
- `PDM_OpenApp()` – apre l'interfaccia PDM

### Bridge Python (`macros/sw_bridge.py`)

Chiamato automaticamente dalla macro VBA:

```powershell
pythonw sw_bridge.py --action checkout --file "C:\...\parte.SLDPRT"
pythonw sw_bridge.py --action checkin  --file "C:\...\parte.SLDPRT"
pythonw sw_bridge.py --action import_props --file "..."
pythonw sw_bridge.py --action export_props --file "..."
pythonw sw_bridge.py --action open
```

---

## Gestione utenti

Ruoli disponibili (da Strumenti → Gestione utenti, solo admin):

| Ruolo | Checkout | Check-in | Crea doc | Rilascia | Admin |
|---|---|---|---|---|---|
| Utente | ✗ | ✗ | ✗ | ✗ | ✗ |
| Progettista | ✓ | ✓ | ✓ | ✗ | ✗ |
| Responsabile | ✓ | ✓ | ✓ | ✓ | ✗ |
| Amministratore | ✓ | ✓ | ✓ | ✓ | ✓ |

**Credenziali default:**
- Username: `admin`
- Password: `admin`

---

## Consigli per ambiente di produzione

1. **Backup**: pianificare backup periodici di `\PDM\database\pdm.db`
2. **Permessi cartella**: tutti gli utenti devono avere R/W sulla cartella condivisa
3. **Antivirus**: escludere la cartella PDM dalla scansione in tempo reale (causa lentezza SQLite)
4. **Rete lenta**: in caso di rete lenta aumentare il timeout in `database.py` (`timeout=10`)
5. **SolidWorks + PDM**: configurare SolidWorks per non bloccare i file con il proprio lock (usare il lock PDM)
6. **Monitoraggio**: usare la vista Workspace per vedere chi ha file in checkout

---

## Estensioni consigliate

Possibili sviluppi futuri:
- **Viste 3D inline**: anteprima modelli con VTK o Open CASCADE
- **PDF automatico**: generazione automatica PDF da disegni alla release
- **Email notifiche**: notifiche email sui cambi di stato
- **Report Excel**: esportazione BOM completa in Excel
- **Integrazione ERP**: export dati verso sistemi gestionali
