# =============================================================================
#  ui/checkin_dialog.py  –  Dialog per check-in (singolo e ASM)
# =============================================================================
import logging
import threading
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QCheckBox, QGroupBox, QTableWidget,
    QTableWidgetItem, QHeaderView, QMessageBox, QTextEdit
)
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QColor

from ui.session import session
from ui.styles import TYPE_ICON
from core.checkout_manager import READONLY_STATES
from config import WORKFLOW_STATES


_WORKER_SCRIPT = (
    __import__('pathlib').Path(__file__).parent.parent / "core" / "thumb_worker.py"
)
_THUMB_TIMEOUT = 40   # secondi massimi per la generazione


def _generate_thumbnail(doc: dict, sp):
    """Genera thumbnail PNG del file archiviato via subprocess isolato (thumb_worker.py)."""
    import sys, subprocess, os, tempfile
    from pathlib import Path

    ext_map = {"Parte": ".SLDPRT", "Assieme": ".SLDASM", "Disegno": ".SLDDRW"}
    file_name = doc.get("file_name") or (
        doc.get("code", "") + ext_map.get(doc.get("doc_type", ""), "")
    )
    try:
        arch_path = sp.archive_path(doc["code"], doc["revision"])
        src_file = arch_path / file_name if file_name else None
    except Exception:
        src_file = None

    if not src_file or not src_file.exists():
        if doc.get("archive_path"):
            src_file = sp.root / Path(doc["archive_path"])
    if not src_file or not src_file.exists():
        logging.warning("Thumbnail: file sorgente non trovato per %s rev %s",
                        doc.get("code"), doc.get("revision"))
        return

    sp.thumbnails.mkdir(parents=True, exist_ok=True)
    suffix = "_DRW" if doc.get("doc_type") == "Disegno" else ""
    dest = sp.thumbnails / f"{doc['code']}_{doc['revision']}{suffix}.png"

    tmp_fd, tmp_log = tempfile.mkstemp(suffix='.txt', prefix='thumb_')
    os.close(tmp_fd)
    try:
        with open(tmp_log, 'w') as lf:
            proc = subprocess.Popen(
                [sys.executable, str(_WORKER_SCRIPT), str(src_file), str(dest)],
                stdout=lf, stderr=lf,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
        proc.wait(timeout=_THUMB_TIMEOUT)
        if dest.exists():
            logging.info("Thumbnail generata: %s", dest)
        else:
            logging.warning("Thumbnail non generata per %s (exit %s)",
                            doc.get("code"), proc.returncode)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.wait()
        logging.warning("Thumbnail timeout per %s", doc.get("code"))
    except Exception as e:
        logging.warning("Thumbnail generation fallita: %s", e)
    finally:
        try:
            os.unlink(tmp_log)
        except OSError:
            pass


class CheckinDialog(QDialog):
    """
    Dialog per check-in.
    - Singolo: mostra info modifica + checkbox elimina da WS
    - ASM: tabella con tutti i file, selezionabili solo quelli in checkout mio
    """

    def __init__(self, document_id: int, parent=None):
        super().__init__(parent)
        self.document_id = document_id
        self.doc = session.files.get_document(document_id)
        self.setWindowTitle("Check-in")
        self.setMinimumSize(600, 400)
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(16, 16, 16, 16)

        if not self.doc:
            layout.addWidget(QLabel("Documento non trovato"))
            return

        self._checkin_blocked = False

        # Info documento principale
        icon = TYPE_ICON.get(self.doc["doc_type"], "")
        lbl_title = QLabel(f"{icon}  {self.doc['code']}  Rev. {self.doc['revision']}  \u2014  {self.doc['title'] or ''}")
        lbl_title.setObjectName("title_label")
        layout.addWidget(lbl_title)

        if self.doc["doc_type"] == "Assieme":
            self._build_bom_table(layout)
            self._build_asm_table(layout)
        else:
            self._build_single_checkin(layout)

        # Opzioni globali
        grp_opt = QGroupBox("Opzioni")
        opt_layout = QVBoxLayout(grp_opt)

        self.chk_delete_ws = QCheckBox("Elimina file dalla workspace dopo il check-in")
        self.chk_delete_ws.setChecked(False)
        opt_layout.addWidget(self.chk_delete_ws)

        # Note
        opt_layout.addWidget(QLabel("Note (opzionale):"))
        self.txt_notes = QTextEdit()
        self.txt_notes.setMaximumHeight(60)
        opt_layout.addWidget(self.txt_notes)

        layout.addWidget(grp_opt)

        # Bottoni
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        btn_cancel = QPushButton("Annulla")
        btn_cancel.clicked.connect(self.reject)
        btn_ok = QPushButton("\U0001f4e5  Check-in")
        btn_ok.setObjectName("btn_primary")
        btn_ok.clicked.connect(self._do_checkin)
        btn_ok.setEnabled(not self._checkin_blocked)
        btn_row.addWidget(btn_cancel)
        btn_row.addWidget(btn_ok)
        layout.addLayout(btn_row)

    # ------------------------------------------------------------------
    def _build_single_checkin(self, layout):
        """UI per checkin singolo con info modifica."""
        mod_info = session.checkout.is_file_modified(self.document_id)

        grp = QGroupBox("Stato file")
        grp_layout = QVBoxLayout(grp)

        if not mod_info["ws_exists"]:
            lbl = QLabel("⛔  File non trovato nella workspace — check-in non disponibile.\n"
                         "Il file deve trovarsi in: " + str(session.checkout._ws_file_path(self.doc)))
            lbl.setWordWrap(True)
            lbl.setStyleSheet("color: #f38ba8; font-weight: bold;")
            grp_layout.addWidget(lbl)
            self._checkin_blocked = True
        elif mod_info["conflict"]:
            lbl = QLabel(
                "\u26a0\ufe0f  CONFLITTO: il file in archivio \u00e8 stato aggiornato "
                "da un altro utente dopo il checkout.\n"
                "Verificare prima di procedere."
            )
            lbl.setWordWrap(True)
            lbl.setStyleSheet("color: #f38ba8; font-weight: bold;")
            grp_layout.addWidget(lbl)
        elif mod_info["modified"]:
            lbl = QLabel("\u2705  File modificato \u2014 verr\u00e0 archiviato")
            lbl.setStyleSheet("color: #a6e3a1;")
            grp_layout.addWidget(lbl)
        else:
            lbl = QLabel(
                "\u26a0\ufe0f  Il file NON risulta modificato.\n"
                "Si consiglia di non archiviare. Si pu\u00f2 rilasciare solo il lock."
            )
            lbl.setWordWrap(True)
            lbl.setStyleSheet("color: #fab387;")
            grp_layout.addWidget(lbl)

            self.chk_archive_anyway = QCheckBox("Archivia comunque")
            self.chk_archive_anyway.setChecked(False)
            grp_layout.addWidget(self.chk_archive_anyway)

        self._single_mod_info = mod_info
        layout.addWidget(grp)

        # DRW companion per Parte/Assieme
        self._drw_companion_id = None
        self.chk_drw = None
        if self.doc["doc_type"] in ("Parte", "Assieme"):
            drw_doc = session.files.get_drw_document(self.document_id)
            grp_drw = QGroupBox("Disegno associato (DRW)")
            drw_layout = QVBoxLayout(grp_drw)
            uid = session.user["id"]
            if drw_doc and drw_doc.get("is_locked") and drw_doc.get("locked_by") == uid:
                drw_mod  = session.checkout.is_file_modified(drw_doc["id"])
                drw_state = drw_doc["state"]
                drw_color = WORKFLOW_STATES.get(drw_state, {}).get("color", "#cdd6f4")
                mod_text  = "✅ modificato" if drw_mod.get("modified") else "⚠️ non modificato"
                self.chk_drw = QCheckBox(
                    f"Esegui check-in DRW  {drw_doc['code']}  — "
                    f"{drw_state}  ({mod_text})"
                )
                self.chk_drw.setStyleSheet(f"color: {drw_color};")
                self.chk_drw.setChecked(drw_mod.get("modified", False))
                self._drw_companion_id = drw_doc["id"]
                drw_layout.addWidget(self.chk_drw)
            elif drw_doc and drw_doc.get("is_locked"):
                locker = drw_doc.get("locked_by_name", "altro utente")
                lbl_drw = QLabel(f"⚠️  DRW in checkout da {locker} — non disponibile")
                lbl_drw.setStyleSheet("color: #f38ba8;")
                drw_layout.addWidget(lbl_drw)
            elif drw_doc:
                lbl_drw = QLabel(f"ℹ️  DRW {drw_doc['code']} — non in checkout")
                lbl_drw.setStyleSheet("color: #6c7086;")
                drw_layout.addWidget(lbl_drw)
            else:
                lbl_drw = QLabel("— Nessun DRW associato")
                lbl_drw.setStyleSheet("color: #6c7086;")
                drw_layout.addWidget(lbl_drw)
            layout.addWidget(grp_drw)

    # ------------------------------------------------------------------
    def _build_bom_table(self, layout):
        """Tabella componenti BOM dell'assieme (solo informativa)."""
        components = session.asm.get_bom_flat(self.document_id)
        if not components:
            return

        grp = QGroupBox(f"Struttura assieme ({len(components)} componenti)")
        grp_layout = QVBoxLayout(grp)

        tbl = QTableWidget(0, 5)
        tbl.setHorizontalHeaderLabels(["Codice", "Tipo", "Rev.", "Titolo", "Qtà"])
        tbl.horizontalHeader().setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch
        )
        tbl.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        tbl.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        tbl.verticalHeader().setVisible(False)
        tbl.setMaximumHeight(160)

        for c in components:
            r = tbl.rowCount()
            tbl.insertRow(r)
            tbl.setItem(r, 0, QTableWidgetItem(c["code"]))
            icon = TYPE_ICON.get(c["doc_type"], "")
            tbl.setItem(r, 1, QTableWidgetItem(f"{icon} {c['doc_type']}"))
            tbl.setItem(r, 2, QTableWidgetItem(c["revision"]))
            tbl.setItem(r, 3, QTableWidgetItem(c.get("title", "")))
            qty = c.get("quantity", 1)
            tbl.setItem(r, 4, QTableWidgetItem(str(int(qty) if qty == int(qty) else qty)))

        grp_layout.addWidget(tbl)
        layout.addWidget(grp)

    # ------------------------------------------------------------------
    def _build_asm_table(self, layout):
        """UI per checkin ASM con tabella selezionabile."""
        grp = QGroupBox("File in workspace per questo assieme")
        grp_layout = QVBoxLayout(grp)

        uid = session.user["id"]
        # Raccogli tutti i file in workspace legati a questo ASM
        ws_files = session.checkout.get_workspace_files(uid)
        # Filtra: il doc principale + componenti con parent_checkout_id
        related = []
        for wf in ws_files:
            if wf["document_id"] == self.document_id:
                related.append(wf)
            elif wf.get("parent_checkout_id") == self.document_id:
                related.append(wf)

        # Aggiungi anche file in checkout diretto che hanno stesso codice (DRW)
        doc_code = self.doc["code"]
        for wf in ws_files:
            if wf["code"] == doc_code and wf["document_id"] != self.document_id:
                if wf not in related:
                    related.append(wf)

        self.tbl_asm = QTableWidget(0, 6)
        self.tbl_asm.setHorizontalHeaderLabels([
            "", "Codice", "Tipo", "Ruolo", "Stato", "Modificato"
        ])
        hdr = self.tbl_asm.horizontalHeader()
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.tbl_asm.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.tbl_asm.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.tbl_asm.verticalHeader().setVisible(False)

        self._asm_checkin_items = []

        for wf in related:
            r = self.tbl_asm.rowCount()
            self.tbl_asm.insertRow(r)

            doc_id = wf["document_id"]
            is_mine = (wf.get("is_locked") and wf.get("locked_by") == uid)
            is_readonly = wf.get("state", "") in READONLY_STATES
            can_checkin = is_mine and not is_readonly

            # Checkbox
            chk = QCheckBox()
            if can_checkin:
                # Controlla se modificato
                mod = session.checkout.is_file_modified(doc_id)
                chk.setChecked(mod.get("modified", False))
                chk.setEnabled(True)
                mod_text = "\u2705 S\u00ec" if mod.get("modified", False) else "\u274c No"
                if mod.get("conflict", False):
                    mod_text = "\u26a0\ufe0f Conflitto"
            else:
                chk.setChecked(False)
                chk.setEnabled(False)
                mod_text = "\u2014"

            self.tbl_asm.setCellWidget(r, 0, chk)
            self.tbl_asm.setItem(r, 1, QTableWidgetItem(wf["code"]))

            icon = TYPE_ICON.get(wf.get("doc_type", ""), "")
            self.tbl_asm.setItem(r, 2, QTableWidgetItem(f"{icon} {wf.get('doc_type', '')}"))

            role_label = {
                "checkout": "Checkout",
                "component": "Componente (copia)",
                "consultation": "Consultazione",
            }.get(wf.get("role", ""), wf.get("role", ""))
            self.tbl_asm.setItem(r, 3, QTableWidgetItem(role_label))

            state_item = QTableWidgetItem(wf.get("state", ""))
            from config import WORKFLOW_STATES
            color = WORKFLOW_STATES.get(wf.get("state", ""), {}).get("color", "#9E9E9E")
            state_item.setForeground(QColor(color))
            self.tbl_asm.setItem(r, 4, state_item)

            mod_item = QTableWidgetItem(mod_text)
            self.tbl_asm.setItem(r, 5, mod_item)

            # Grayed out per non selezionabili
            if not can_checkin:
                for col in range(1, 6):
                    item = self.tbl_asm.item(r, col)
                    if item:
                        item.setForeground(QColor("#6c7086"))

            self._asm_checkin_items.append({
                "document_id": doc_id,
                "checkbox": chk,
                "can_checkin": can_checkin,
            })

        grp_layout.addWidget(self.tbl_asm)
        layout.addWidget(grp)

    # ------------------------------------------------------------------
    def _do_checkin(self):
        try:
            notes = self.txt_notes.toPlainText().strip()
            delete_ws = self.chk_delete_ws.isChecked()

            if self.doc["doc_type"] == "Assieme":
                self._do_checkin_asm(notes, delete_ws)
            else:
                self._do_checkin_single(notes, delete_ws)

        except Exception as e:
            QMessageBox.critical(self, "Errore check-in", str(e))

    def _sync_sw_to_pdm_before_checkin(self, document_id: int):
        """Importa le proprietà SW nel PDM dal file in workspace prima dell'archiviazione."""
        try:
            ws_file = session.checkout._ws_file_path(
                session.files.get_document(document_id)
            )
            if ws_file and ws_file.exists():
                session.properties.sync_sw_to_pdm(document_id, ws_file)
        except Exception as e:
            logging.warning("Sync SW→PDM pre-checkin fallita per doc %s: %s", document_id, e)

    def _do_checkin_single(self, notes: str, delete_ws: bool):
        mod_info = getattr(self, "_single_mod_info", {})
        archive = True

        if not mod_info.get("modified", True):
            # File non modificato
            archive_anyway = getattr(self, "chk_archive_anyway", None)
            if archive_anyway and archive_anyway.isChecked():
                archive = True
            else:
                archive = False

        # Importa proprietà SW nel PDM prima di archiviare
        if archive:
            self._sync_sw_to_pdm_before_checkin(self.document_id)

        result = session.checkout.checkin(
            self.document_id,
            archive_file=archive,
            delete_from_workspace=delete_ws,
            notes=notes,
        )

        # Genera thumbnail in background se il file è stato archiviato
        if result.get("archived") and session.sp:
            import threading
            doc_snap = session.files.get_document(self.document_id)
            sp_snap = session.sp
            threading.Thread(
                target=_generate_thumbnail,
                args=(doc_snap, sp_snap),
                daemon=True,
            ).start()

        if result.get("conflict"):
            r = QMessageBox.warning(
                self, "Conflitto rilevato",
                "Il file in archivio \u00e8 stato modificato da un altro utente.\n"
                "Il check-in \u00e8 stato eseguito comunque.\n"
                "Verificare il file archiviato.",
                QMessageBox.StandardButton.Ok,
            )

        msg = "Check-in eseguito."
        if result.get("archived"):
            msg += "\nFile archiviato."
        else:
            msg += "\nLock rilasciato (file non archiviato)."

        # Checkin DRW companion se selezionato
        drw_id = getattr(self, "_drw_companion_id", None)
        chk_drw = getattr(self, "chk_drw", None)
        if drw_id and chk_drw and chk_drw.isChecked():
            try:
                drw_mod = session.checkout.is_file_modified(drw_id)
                r_drw = session.checkout.checkin(
                    drw_id,
                    archive_file=drw_mod.get("modified", True),
                    delete_from_workspace=delete_ws,
                    notes=notes,
                )
                msg += "\nDRW archiviato." if r_drw.get("archived") else "\nDRW: lock rilasciato."
            except Exception as e_drw:
                msg += f"\n\n⚠️ Errore DRW: {e_drw}"

        QMessageBox.information(self, "OK", msg)
        self.accept()

    def _do_checkin_asm(self, notes: str, delete_ws: bool):
        """Checkin di tutti i file selezionati nella tabella ASM."""
        checked_ids = []
        for item in self._asm_checkin_items:
            if item["can_checkin"] and item["checkbox"].isChecked():
                checked_ids.append(item["document_id"])

        if not checked_ids:
            QMessageBox.warning(
                self, "Nessun file selezionato",
                "Selezionare almeno un file per il check-in.\n"
                "Oppure annullare per uscire."
            )
            return

        results = []
        errors = []
        for doc_id in checked_ids:
            try:
                # Importa proprietà SW nel PDM prima di archiviare
                self._sync_sw_to_pdm_before_checkin(doc_id)
                r = session.checkout.checkin(
                    doc_id,
                    archive_file=True,
                    delete_from_workspace=delete_ws,
                    notes=notes,
                )
                results.append(r)
                # Thumbnail in background per ogni file archiviato
                if r.get("archived") and session.sp:
                    doc_snap = session.files.get_document(doc_id)
                    sp_snap = session.sp
                    threading.Thread(
                        target=_generate_thumbnail,
                        args=(doc_snap, sp_snap),
                        daemon=True,
                    ).start()
            except Exception as e:
                errors.append(f"Errore su doc {doc_id}: {e}")

        msg = f"Check-in completato: {len(results)} file archiviati."
        if errors:
            msg += f"\n\n\u26a0\ufe0f  Errori ({len(errors)}):\n" + "\n".join(errors)

        conflicts = [r for r in results if r.get("conflict")]
        if conflicts:
            msg += f"\n\n\u26a0\ufe0f  {len(conflicts)} file con conflitto (verificare)."

        QMessageBox.information(self, "Check-in ASM", msg)
        self.accept()
