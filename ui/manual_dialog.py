# =============================================================================
#  ui/manual_dialog.py  –  Manuale utente PDM-SW
# =============================================================================
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QTabWidget, QWidget,
    QTextBrowser, QPushButton, QLabel, QFrame
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QFont

from config import APP_NAME, APP_VERSION


# ---------------------------------------------------------------------------
# Contenuto del manuale (HTML per tab)
# ---------------------------------------------------------------------------

_CSS = """
<style>
body  { font-family: Segoe UI, Arial, sans-serif; font-size: 13px; color: #cdd6f4; }
h1    { color: #89b4fa; border-bottom: 2px solid #313244; padding-bottom: 6px; }
h2    { color: #89dceb; margin-top: 20px; }
h3    { color: #a6e3a1; margin-top: 14px; }
code  { background: #313244; padding: 2px 6px; border-radius: 4px;
        font-family: Consolas, monospace; color: #f38ba8; }
pre   { background: #1e1e2e; padding: 10px; border-radius: 6px;
        font-family: Consolas, monospace; font-size: 12px;
        color: #cba6f7; border-left: 3px solid #89b4fa; }
table { border-collapse: collapse; width: 100%; margin: 10px 0; }
th    { background: #313244; color: #89b4fa; padding: 6px 10px;
        text-align: left; border-bottom: 2px solid #45475a; }
td    { padding: 5px 10px; border-bottom: 1px solid #313244; }
tr:nth-child(even) { background: #181825; }
.warn { background: #45350a; border-left: 4px solid #f9e2af;
        padding: 8px 12px; border-radius: 4px; color: #f9e2af; }
.tip  { background: #0a3020; border-left: 4px solid #a6e3a1;
        padding: 8px 12px; border-radius: 4px; color: #a6e3a1; }
.note { background: #1a1a35; border-left: 4px solid #89b4fa;
        padding: 8px 12px; border-radius: 4px; color: #89b4fa; }
ul li { margin: 4px 0; }
ol li { margin: 4px 0; }
hr    { border: none; border-top: 1px solid #313244; margin: 16px 0; }
</style>
"""

# ── 1. INSTALLAZIONE ─────────────────────────────────────────────────────────
_TAB_INSTALLAZIONE = _CSS + """
<h1>Installazione</h1>

<h2>Requisiti di sistema</h2>
<table>
  <tr><th>Componente</th><th>Requisito</th></tr>
  <tr><td>Sistema operativo</td><td>Windows 10 / 11 (64-bit)</td></tr>
  <tr><td>Python</td><td>3.11 o superiore</td></tr>
  <tr><td>SolidWorks</td><td>2020 o superiore (opzionale, per integrazione COM)</td></tr>
  <tr><td>Rete</td><td>Cartella condivisa accessibile da tutti i client (UNC o lettera di unità)</td></tr>
  <tr><td>RAM</td><td>4 GB minimi, 8 GB consigliati</td></tr>
  <tr><td>Spazio disco</td><td>200 MB per l'applicazione + spazio archivio CAD</td></tr>
</table>

<h2>Struttura cartelle</h2>
<pre>
PDM-SW/                      ← cartella applicazione (può essere in rete)
  ├── main.py                ← punto di avvio
  ├── config.py
  ├── core/                  ← logica business
  ├── ui/                    ← interfaccia grafica
  ├── macros/                ← macro SolidWorks
  ├── .venv/                 ← ambiente virtuale Python
  └── requirements.txt

ARCHIVIO_PDM/                ← cartella condivisa in rete
  ├── pdm_database.sqlite    ← database condiviso
  ├── archive/               ← file CAD archiviati
  │     ├── Parte/
  │     ├── Assieme/
  │     └── Disegno/
  └── backups/               ← backup automatici
</pre>

<h2>Procedura di installazione</h2>
<ol>
  <li>Copiare la cartella <code>PDM-SW</code> sul PC (locale o in rete).</li>
  <li>Creare l'ambiente virtuale:<br>
    <pre>python -m venv .venv
.venv\\Scripts\\activate
pip install -r requirements.txt</pre>
  </li>
  <li>Avviare l'applicazione:<br>
    <pre>python main.py</pre>
    oppure fare doppio clic su <code>avvia_pdm.bat</code>.
  </li>
  <li>Al primo avvio comparirà il wizard di configurazione (percorso archivio di rete).</li>
</ol>

<h2>Installazione su più workstation</h2>
<ol>
  <li>Ripetere i passi 1-3 su ogni PC client.</li>
  <li>Tutti i client puntano alla <em>stessa</em> cartella condivisa in rete.</li>
  <li>Il database SQLite è condiviso: massimo <strong>5 utenti simultanei</strong> consigliati.</li>
</ol>

<div class="tip">
  <b>Suggerimento:</b> Creare un collegamento sul Desktop a <code>avvia_pdm.bat</code>
  per un accesso rapido.
</div>

<h2>Aggiornamento</h2>
<ol>
  <li>Sostituire i file dell'applicazione con la versione nuova.</li>
  <li>Eseguire: <pre>pip install -r requirements.txt --upgrade</pre></li>
  <li>Il database si aggiorna automaticamente alla prima connessione (migration).</li>
</ol>
"""

# ── 2. CONFIGURAZIONE ────────────────────────────────────────────────────────
_TAB_CONFIGURAZIONE = _CSS + """
<h1>Configurazione</h1>

<h2>Configurazione percorso rete</h2>
<p>
  <b>Strumenti → Configurazione percorso rete…</b><br>
  Imposta la cartella condivisa in rete che conterrà il database e l'archivio CAD.
  Ogni workstation deve puntare allo stesso percorso.
</p>
<table>
  <tr><th>Campo</th><th>Descrizione</th></tr>
  <tr><td>Percorso archivio</td><td>UNC (\\\\server\\PDM) o lettera di unità (Z:\\PDM)</td></tr>
</table>

<h2>Configurazione SolidWorks</h2>
<p>
  <b>Strumenti → Configurazione SolidWorks…</b><br>
  Configurazione locale per ogni workstation (non condivisa).
</p>
<table>
  <tr><th>Campo</th><th>Descrizione</th></tr>
  <tr><td>Workspace locale</td><td>Cartella locale dove vengono copiati i file in checkout</td></tr>
  <tr><td>Template Parte (.SLDPRT)</td><td>File template usato per creare nuove parti</td></tr>
  <tr><td>Template Assieme (.SLDASM)</td><td>File template per nuovi assiemi</td></tr>
  <tr><td>Template Disegno (.SLDDRW)</td><td>File template per nuovi disegni</td></tr>
  <tr><td>Proprietà da mappare</td><td>Mapping tra proprietà PDM e proprietà custom SolidWorks</td></tr>
</table>

<h2>Schema di Codifica</h2>
<p>
  <b>Strumenti → Schema di Codifica…</b> (solo Amministratori)<br>
  Definisce il formato dei codici per ogni livello della gerarchia.
</p>
<table>
  <tr><th>Impostazione</th><th>Descrizione</th></tr>
  <tr><td>Codice Macchina</td><td>Tipo (ALPHA/NUM) e lunghezza (es. ALPHA, 3 car. → "ABC")</td></tr>
  <tr><td>Codice Gruppo</td><td>Tipo e lunghezza del codice gruppo</td></tr>
  <tr><td>Template LIV0</td><td>Formato codice Macchina ASM (es. <code>{MACH}_V{VER:3}</code>)</td></tr>
  <tr><td>Template LIV1</td><td>Formato codice Gruppo ASM (es. <code>{MACH}_{GRP}-V{VER:3}</code>)</td></tr>
  <tr><td>Template LIV2/1</td><td>Formato Sottogruppo ASM con range numerazione</td></tr>
  <tr><td>Template LIV2/2</td><td>Formato Parte PRT con range numerazione</td></tr>
  <tr><td>Warning collisione</td><td>Soglia avviso quando i due contatori LIV2 si avvicinano</td></tr>
</table>

<h3>Variabili template</h3>
<table>
  <tr><th>Variabile</th><th>Significato</th><th>Esempio</th></tr>
  <tr><td><code>{MACH}</code></td><td>Codice macchina</td><td>ABC</td></tr>
  <tr><td><code>{GRP}</code></td><td>Codice gruppo</td><td>COMP</td></tr>
  <tr><td><code>{VER:N}</code></td><td>Versione, N cifre</td><td>{VER:3} → 001</td></tr>
  <tr><td><code>{NUM:N}</code></td><td>Numero sequenziale, N cifre</td><td>{NUM:4} → 0001</td></tr>
</table>

<div class="warn">
  <b>Attenzione:</b> Le modifiche allo schema si applicano solo ai nuovi codici.
  I codici già presenti nel database non vengono modificati.
</div>

<h2>Gestione utenti</h2>
<p><b>Strumenti → Gestione utenti…</b> (solo Amministratori)</p>
<table>
  <tr><th>Ruolo</th><th>Permessi</th></tr>
  <tr><td><b>Admin</b></td><td>Accesso completo: gestione utenti, schema, workflow avanzato, reset contatori</td></tr>
  <tr><td><b>Responsabile</b></td><td>Crea documenti, gestisce workflow, approva revisioni</td></tr>
  <tr><td><b>Progettista</b></td><td>Crea documenti, checkout/check-in, modifica in lavorazione</td></tr>
  <tr><td><b>Visualizzatore</b></td><td>Solo lettura dell'archivio</td></tr>
</table>

<h2>Profili PDM</h2>
<p>
  <b>Strumenti → PDM Profile…</b><br>
  Permette di gestire più configurazioni di database (es. Produzione, Test).
  Il profilo attivo è mostrato nella barra di stato.
</p>

<h2>Macro SolidWorks</h2>
<p>
  Il file <code>macros/PDM_Integration.swb</code> va caricato in SolidWorks
  e assegnato a un pulsante nella barra degli strumenti.<br>
  <b>SolidWorks → Strumenti → Personalizza → Macro</b>
</p>
"""

# ── 3. CODIFICA ──────────────────────────────────────────────────────────────
_TAB_CODIFICA = _CSS + """
<h1>Schema di Codifica Gerarchico</h1>

<h2>Struttura gerarchica</h2>
<p>Ogni documento è identificato da un codice univoco generato automaticamente
secondo la gerarchia:</p>

<table>
  <tr><th>Livello</th><th>Nome</th><th>Tipo doc</th><th>Esempio codice</th></tr>
  <tr><td>LIV0</td><td>Macchina</td><td>Assieme</td><td>ABC_V001</td></tr>
  <tr><td>LIV1</td><td>Gruppo</td><td>Assieme</td><td>ABC_COMP-V001</td></tr>
  <tr><td>LIV2/1</td><td>Sottogruppo</td><td>Assieme</td><td>ABC_COMP-9999</td></tr>
  <tr><td>LIV2/2</td><td>Parte</td><td>Parte</td><td>ABC_COMP-0001</td></tr>
</table>

<h2>Macchine e Gruppi</h2>
<p>
  <b>Strumenti → Macchine e Gruppi…</b><br>
  Prima di creare documenti è necessario definire le macchine e i gruppi.
</p>
<ul>
  <li><b>Macchina</b> (LIV0): entità di primo livello, es. "ABC – Linea di assemblaggio"</li>
  <li><b>Gruppo</b> (LIV1): sotto-componente di una macchina, es. "COMP – Compressore"</li>
</ul>

<h3>Generazione massiva</h3>
<p>
  Il pulsante <b>"Importa / Genera serie…"</b> nei tab Macchine e Gruppi permette di:
</p>
<ul>
  <li><b>Genera serie</b>: definire primo e ultimo codice per generare automaticamente
    tutta la sequenza (ALPHA: AAA→AZZ; NUM: 001→050)</li>
  <li><b>Lista libera</b>: incollare righe <code>CODICE ; Descrizione</code></li>
</ul>

<h2>Contatori</h2>
<p>
  I contatori vengono incrementati automaticamente ad ogni creazione documento.
</p>
<table>
  <tr><th>Contatore</th><th>Per livello</th><th>Direzione default</th></tr>
  <tr><td>VERSION</td><td>LIV0 e LIV1</td><td>Ascendente (001, 002, ...)</td></tr>
  <tr><td>PART</td><td>LIV2/1 Sottogruppo</td><td>Configurabile</td></tr>
  <tr><td>SUBGROUP</td><td>LIV2/2 Parte</td><td>Configurabile</td></tr>
</table>

<div class="note">
  <b>Collisione LIV2:</b> I contatori PART e SUBGROUP condividono lo stesso spazio
  numerico. Il sistema avvisa quando la distanza residua scende sotto la soglia configurata
  (default: 500 codici).
</div>

<h2>Creazione di un nuovo documento</h2>
<ol>
  <li>Cliccare <b>File → Nuovo documento…</b> (Ctrl+N) o <b>➕ Nuovo</b> nella toolbar.</li>
  <li>Selezionare <b>Macchina</b> e <b>Gruppo</b> dall'elenco.</li>
  <li>Scegliere il <b>Livello</b>:
    <ul>
      <li>LIV0 – Macchina (non richiede gruppo)</li>
      <li>LIV1 – Gruppo</li>
      <li>LIV2/1 – Sottogruppo (Assieme)</li>
      <li>LIV2/2 – Parte</li>
    </ul>
  </li>
  <li>Il <b>Codice generato</b> viene mostrato in anteprima in tempo reale.</li>
  <li>Inserire il <b>Titolo</b> (obbligatorio) e la <b>Revisione</b>.</li>
  <li>Scegliere la modalità di creazione:
    <ul>
      <li><b>Solo codice</b>: registra il codice senza creare file SW</li>
      <li><b>Crea documento</b>: apre il template in SolidWorks e salva il file</li>
      <li><b>Crea documenti (+ DRW)</b>: crea anche il file disegno</li>
    </ul>
  </li>
  <li>Cliccare <b>Crea Documento</b>.</li>
</ol>

<div class="tip">
  Dopo la creazione, il form si azzera automaticamente e mostra il prossimo
  codice disponibile, permettendo di creare documenti in sequenza rapidamente.
</div>
"""

# ── 4. FUNZIONAMENTO PDM ─────────────────────────────────────────────────────
_TAB_PDM = _CSS + """
<h1>Funzionamento del PDM</h1>

<h2>Archivio CAD</h2>
<p>
  Il tab <b>🗄️ Archivio CAD</b> mostra tutti i documenti codificati.
  Ogni riga rappresenta un documento con:
</p>
<table>
  <tr><th>Colonna</th><th>Descrizione</th></tr>
  <tr><td>Codice</td><td>Codice PDM univoco del documento</td></tr>
  <tr><td>Rev.</td><td>Revisione corrente</td></tr>
  <tr><td>Tipo</td><td>Parte / Assieme / Disegno</td></tr>
  <tr><td>Titolo</td><td>Descrizione del documento</td></tr>
  <tr><td>Stato</td><td>Stato workflow (vedi sezione Workflow)</td></tr>
  <tr><td>Checkout</td><td>Utente/workstation che ha il file in lavorazione</td></tr>
  <tr><td>Data</td><td>Data ultima modifica</td></tr>
</table>

<h2>Non codificati</h2>
<p>
  Il tab <b>🗂️ Non codificati</b> mostra i file presenti nella cartella archivio
  ma non ancora registrati nel database. Utile per importare file esistenti.
</p>

<h2>Workspace</h2>
<p>
  Il tab <b>📁 Workspace</b> mostra i file presenti nella workspace locale
  del PC corrente. Permette di:
</p>
<ul>
  <li>Vedere quali file sono in checkout locale</li>
  <li>Fare check-in direttamente dalla lista</li>
  <li>Aprire file in SolidWorks</li>
</ul>

<h2>Dettaglio documento</h2>
<p>Fare doppio clic su un documento per aprire il pannello dettaglio, che contiene:</p>
<ul>
  <li><b>Generale</b>: codice, revisione, tipo, titolo, stato, checkout, file associato</li>
  <li><b>Proprietà SW</b>: proprietà custom SolidWorks (importabili/esportabili)</li>
  <li><b>Struttura (BOM)</b>: distinta base gerarchica</li>
  <li><b>Storico</b>: log di tutte le operazioni sul documento</li>
</ul>

<h2>Gestione file</h2>
<p>I pulsanti nel pannello dettaglio permettono di:</p>
<table>
  <tr><th>Pulsante</th><th>Azione</th></tr>
  <tr><td>✨ Crea in SW</td><td>Crea file da template in workspace e apre in SolidWorks</td></tr>
  <tr><td>📋 Crea da file</td><td>Copia file esistente in workspace rinominandolo col codice PDM</td></tr>
  <tr><td>📥 Importa in PDM</td><td>Archivia il file dalla workspace nell'archivio condiviso</td></tr>
  <tr><td>📤 Esporta in WS</td><td>Copia il file dall'archivio alla workspace locale</td></tr>
  <tr><td>🔧 Apri in SW</td><td>Copia dall'archivio in workspace e apre in SolidWorks</td></tr>
</table>

<h2>Backup</h2>
<p>
  <b>Strumenti → Backup database…</b><br>
  Crea una copia del database nella cartella <code>backups/</code>.
  I backup sono conservati automaticamente (ultimi N, configurabile).
</p>
<div class="tip">
  Il backup automatico viene eseguito anche prima di ogni migrazione del database.
</div>
"""

# ── 5. WORKFLOW ──────────────────────────────────────────────────────────────
_TAB_WORKFLOW = _CSS + """
<h1>Workflow e Stati documento</h1>

<h2>Ciclo di vita del documento</h2>
<pre>
  ┌─────────────────────────────────────────────────────┐
  │                                                     │
  │   [Creazione] ──→ In Lavorazione ──→ In Revisione  │
  │                        ↑                  ↓        │
  │                        └──── (rimanda) ───┤        │
  │                                           ↓        │
  │                                       Rilasciato   │
  │                                           ↓        │
  │                                        Obsoleto    │
  └─────────────────────────────────────────────────────┘
</pre>

<h2>Stati</h2>
<table>
  <tr><th>Stato</th><th>Icona</th><th>Descrizione</th></tr>
  <tr><td><b>In Lavorazione</b></td><td>🔵</td><td>Documento in fase di progettazione, modificabile</td></tr>
  <tr><td><b>In Revisione</b></td><td>🟡</td><td>Documento in attesa di approvazione, non modificabile</td></tr>
  <tr><td><b>Rilasciato</b></td><td>🟢</td><td>Documento approvato e congelato, non modificabile</td></tr>
  <tr><td><b>Obsoleto</b></td><td>⚫</td><td>Documento ritirato dalla produzione</td></tr>
</table>

<h2>Transizioni consentite</h2>
<table>
  <tr><th>Da</th><th>A</th><th>Chi può</th></tr>
  <tr><td>In Lavorazione</td><td>In Revisione</td><td>Progettista, Responsabile, Admin</td></tr>
  <tr><td>In Revisione</td><td>Rilasciato</td><td>Responsabile, Admin</td></tr>
  <tr><td>In Revisione</td><td>In Lavorazione</td><td>Responsabile, Admin (rimanda)</td></tr>
  <tr><td>Rilasciato</td><td>Obsoleto</td><td>Responsabile, Admin</td></tr>
  <tr><td>Qualsiasi</td><td>In Lavorazione</td><td>Admin (reset forzato)</td></tr>
</table>

<h2>Come cambiare stato</h2>
<ol>
  <li>Selezionare il documento nell'archivio.</li>
  <li>Cliccare <b>🔄 Workflow</b> nella toolbar, oppure fare doppio clic sul documento
    e usare il pulsante Workflow nel pannello dettaglio.</li>
  <li>Scegliere la transizione disponibile e confermare.</li>
</ol>

<div class="warn">
  <b>Nota:</b> Un documento può essere modificato (checkout) solo quando è in stato
  <b>In Lavorazione</b>. Se è In Revisione o Rilasciato, è necessario prima
  riportarlo In Lavorazione (solo Responsabile/Admin).
</div>

<h2>Documento collegato (Companion)</h2>
<p>
  Ogni documento Parte/Assieme può avere un Disegno (.SLDDRW) collegato con lo stesso codice.
  Il pannello dettaglio mostra lo stato del companion e permette di crearlo se mancante.
  Le transizioni di stato vengono sincronizzate automaticamente tra il documento e il suo DRW.
</p>

<h2>Revisioni</h2>
<p>
  La revisione (es. <code>00</code>, <code>01</code>, <code>A</code>) è parte del documento.
  Per creare una nuova revisione di un documento esistente, creare un nuovo documento
  con lo stesso codice e la revisione successiva.
</p>
"""

# ── 6. CHECKOUT / CHECK-IN ───────────────────────────────────────────────────
_TAB_CHECKOUT = _CSS + """
<h1>Check-Out e Check-In</h1>

<h2>Cos'è il Checkout</h2>
<p>
  Il <b>checkout</b> è il meccanismo di blocco che impedisce a più utenti di
  modificare lo stesso file contemporaneamente. Quando un documento è in checkout:
</p>
<ul>
  <li>Il file viene copiato dalla workspace nella cartella locale dell'utente</li>
  <li>Il documento viene <b>bloccato</b>: nessun altro può fare checkout</li>
  <li>L'archivio mostra il nome dell'utente e la workstation che ha il file</li>
</ul>

<h2>Checkout dall'interfaccia PDM</h2>
<ol>
  <li>Selezionare il documento nell'archivio.</li>
  <li>Cliccare <b>📤 Checkout</b> nella toolbar, oppure aprire il documento
    e cliccare <b>📤 Esporta in WS</b>.</li>
  <li>Il file viene copiato nella workspace locale configurata.</li>
  <li>Aprire il file in SolidWorks, modificarlo e salvarlo.</li>
</ol>

<h2>Checkout dalla macro SolidWorks</h2>
<ol>
  <li>Aprire il file direttamente in SolidWorks (da qualsiasi percorso).</li>
  <li>Eseguire la macro <b>PDM_Integration.swb</b>.</li>
  <li>La macro rileva il file aperto, trova il documento nel PDM e
    esegue automaticamente il checkout.</li>
</ol>

<h2>Check-in dall'interfaccia PDM</h2>
<ol>
  <li>Salvare il file in SolidWorks.</li>
  <li>Nell'archivio PDM, selezionare il documento.</li>
  <li>Cliccare <b>📥 Check-in</b> nella toolbar.</li>
  <li>Il file viene copiato dall'workspace locale nell'archivio condiviso.</li>
  <li>Il lock viene rilasciato.</li>
</ol>

<h2>Check-in dalla macro SolidWorks</h2>
<ol>
  <li>Salvare il file in SolidWorks (Ctrl+S).</li>
  <li>Eseguire la macro <b>PDM_Integration.swb</b>.</li>
  <li>La macro esegue automaticamente il check-in.</li>
</ol>

<h2>Rilascio lock forzato</h2>
<p>
  Solo gli <b>Amministratori</b> possono rilasciare un lock senza fare check-in,
  ad esempio se un utente non è disponibile. Aprire il documento → il lock
  compare nella sezione Checkout → usare il menu contestuale per rilasciarlo.
</p>

<div class="warn">
  <b>Attenzione:</b> Il rilascio forzato del lock non riporta le modifiche locali
  nell'archivio. Le modifiche non salvate del file andranno perse.
</div>

<h2>Importazione struttura ASM da SolidWorks</h2>
<p>
  <b>SolidWorks → 📦 Importa struttura ASM da SW…</b><br>
  Wizard che legge la struttura di un assieme aperto in SolidWorks e:
</p>
<ul>
  <li>Crea automaticamente i documenti mancanti nel PDM</li>
  <li>Assegna i codici secondo la gerarchia configurata</li>
  <li>Copia i file nella workspace locale</li>
  <li>Registra la distinta base (BOM)</li>
</ul>
"""

# ── 7. COMANDI ───────────────────────────────────────────────────────────────
_TAB_COMANDI = _CSS + """
<h1>Riferimento Comandi</h1>

<h2>Menu File</h2>
<table>
  <tr><th>Comando</th><th>Scorciatoia</th><th>Descrizione</th></tr>
  <tr><td>Nuovo documento…</td><td>Ctrl+N</td><td>Apre il wizard creazione documento</td></tr>
  <tr><td>Aggiorna archivio</td><td>F5</td><td>Ricarica tutti i tab dall'archivio</td></tr>
  <tr><td>Cambia utente…</td><td>—</td><td>Effettua logout e mostra il login</td></tr>
  <tr><td>Esci</td><td>Alt+F4</td><td>Chiude l'applicazione</td></tr>
</table>

<h2>Menu Strumenti</h2>
<table>
  <tr><th>Comando</th><th>Chi può</th><th>Descrizione</th></tr>
  <tr><td>Schema di Codifica…</td><td>Admin</td><td>Configura format codici e template</td></tr>
  <tr><td>Macchine e Gruppi…</td><td>Admin, Responsabile</td><td>Gestisce entità gerarchiche e contatori</td></tr>
  <tr><td>Gestione utenti…</td><td>Admin</td><td>Crea/modifica utenti e ruoli</td></tr>
  <tr><td>Configurazione SolidWorks…</td><td>Tutti</td><td>Workspace locale e template (per workstation)</td></tr>
  <tr><td>PDM Profile…</td><td>Tutti</td><td>Cambia profilo/database attivo</td></tr>
  <tr><td>Configurazione percorso rete…</td><td>Tutti</td><td>Imposta il percorso archivio condiviso</td></tr>
  <tr><td>Backup database…</td><td>Admin, Responsabile</td><td>Crea backup manuale del database</td></tr>
  <tr><td>Rigenera anteprime…</td><td>Admin</td><td>Rigenera thumbnail via eDrawings</td></tr>
</table>

<h2>Menu SolidWorks</h2>
<table>
  <tr><th>Comando</th><th>Descrizione</th></tr>
  <tr><td>Checkout da SW…</td><td>Istruzioni per eseguire checkout tramite macro</td></tr>
  <tr><td>Check-in da SW…</td><td>Istruzioni per eseguire check-in tramite macro</td></tr>
  <tr><td>📦 Importa struttura ASM da SW…</td><td>Wizard importazione massiva assieme</td></tr>
  <tr><td>Informazioni macro SW</td><td>Mostra percorso e istruzioni installazione macro</td></tr>
</table>

<h2>Toolbar</h2>
<table>
  <tr><th>Pulsante</th><th>Descrizione</th></tr>
  <tr><td>➕ Nuovo</td><td>Crea nuovo documento</td></tr>
  <tr><td>📤 Checkout</td><td>Esegue checkout del documento selezionato</td></tr>
  <tr><td>📥 Check-in</td><td>Esegue check-in del documento selezionato</td></tr>
  <tr><td>🔄 Workflow</td><td>Apre il dialog cambio stato workflow</td></tr>
  <tr><td>↻ Aggiorna</td><td>Ricarica tutti i dati dall'archivio</td></tr>
</table>

<h2>Pannello documento – Tab Generale</h2>
<table>
  <tr><th>Pulsante</th><th>Descrizione</th></tr>
  <tr><td>✨ Crea in SW</td><td>Crea file da template in SolidWorks</td></tr>
  <tr><td>📋 Crea da file</td><td>Copia e rinomina file esistente</td></tr>
  <tr><td>📥 Importa in PDM</td><td>Archivia file dalla workspace</td></tr>
  <tr><td>📤 Esporta in WS</td><td>Copia file nell'archivio alla workspace locale</td></tr>
  <tr><td>🔧 Apri in SW</td><td>Apre il file in SolidWorks</td></tr>
  <tr><td>＋ Crea DRW</td><td>Crea il disegno collegato</td></tr>
  <tr><td>Salva Modifiche</td><td>Salva titolo e descrizione del documento</td></tr>
</table>

<h2>Pannello documento – Tab Proprietà SW</h2>
<table>
  <tr><th>Pulsante</th><th>Descrizione</th></tr>
  <tr><td>+ Aggiungi</td><td>Aggiunge riga proprietà</td></tr>
  <tr><td>- Rimuovi</td><td>Rimuove la riga selezionata</td></tr>
  <tr><td>Importa da SW</td><td>Legge le proprietà custom dal file SolidWorks aperto</td></tr>
  <tr><td>Esporta Excel</td><td>Esporta le proprietà in formato .xlsx</td></tr>
  <tr><td>Importa Excel</td><td>Importa proprietà da file .xlsx</td></tr>
  <tr><td>Salva Proprietà</td><td>Salva nel database le proprietà mostrate</td></tr>
</table>

<h2>Pannello documento – Tab BOM</h2>
<table>
  <tr><th>Pulsante</th><th>Descrizione</th></tr>
  <tr><td>+ Aggiungi componente</td><td>Collega un documento figlio a questo assieme</td></tr>
  <tr><td>- Rimuovi</td><td>Rimuove il collegamento al componente selezionato</td></tr>
  <tr><td>Importa struttura da SW</td><td>Importa la BOM dall'assieme aperto in SolidWorks</td></tr>
  <tr><td>📊 Esporta BOM Excel</td><td>Esporta la distinta base in formato .xlsx</td></tr>
</table>

<h2>Scorciatoie da tastiera</h2>
<table>
  <tr><th>Scorciatoia</th><th>Azione</th></tr>
  <tr><td>Ctrl+N</td><td>Nuovo documento</td></tr>
  <tr><td>F5</td><td>Aggiorna archivio</td></tr>
  <tr><td>Alt+F4</td><td>Esci</td></tr>
  <tr><td>Invio / doppio clic</td><td>Apri documento selezionato</td></tr>
  <tr><td>F3 / Ctrl+F</td><td>Attiva la ricerca nell'archivio (se disponibile)</td></tr>
</table>
"""


# ---------------------------------------------------------------------------
# Dialog
# ---------------------------------------------------------------------------

class ManualDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Manuale utente — {APP_NAME} v{APP_VERSION}")
        self.setMinimumSize(900, 680)
        self._build_ui()

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 8)
        root.setSpacing(0)

        # Header
        hdr = QWidget()
        hdr.setStyleSheet("background: #1e1e2e; padding: 10px 16px;")
        hdr_lyt = QHBoxLayout(hdr)
        hdr_lyt.setContentsMargins(16, 10, 16, 10)
        lbl = QLabel(f"<b>{APP_NAME}</b> — Manuale utente   <small>v{APP_VERSION}</small>")
        lbl.setStyleSheet("color: #cdd6f4; font-size: 14px;")
        hdr_lyt.addWidget(lbl)
        hdr_lyt.addStretch()
        root.addWidget(hdr)

        # Tabs
        tabs = QTabWidget()
        tabs.setDocumentMode(True)

        for label, html in [
            ("📦 Installazione",   _TAB_INSTALLAZIONE),
            ("⚙️ Configurazione",  _TAB_CONFIGURAZIONE),
            ("🔢 Codifica",        _TAB_CODIFICA),
            ("🗄️ PDM",            _TAB_PDM),
            ("🔄 Workflow",        _TAB_WORKFLOW),
            ("📤📥 Checkout/CI",   _TAB_CHECKOUT),
            ("📋 Comandi",         _TAB_COMANDI),
        ]:
            browser = QTextBrowser()
            browser.setOpenExternalLinks(True)
            browser.setStyleSheet(
                "background: #181825; color: #cdd6f4; "
                "border: none; padding: 4px;"
            )
            browser.setHtml(html)
            tabs.addTab(browser, label)

        root.addWidget(tabs, stretch=1)

        # Footer
        btn_row = QHBoxLayout()
        btn_row.setContentsMargins(16, 0, 16, 0)
        btn_row.addStretch()
        btn_close = QPushButton("Chiudi")
        btn_close.setMinimumWidth(100)
        btn_close.clicked.connect(self.accept)
        btn_row.addWidget(btn_close)
        root.addLayout(btn_row)
