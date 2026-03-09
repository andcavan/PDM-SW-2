# =============================================================================
#  core/pdf_worker.py  –  Subprocess: esporta DRW → PDF via SolidWorks COM
#  Uso: python pdf_worker.py <src_drw_path> <dest_pdf_path>
#  Exit 0 = OK, Exit 1 = errore
# =============================================================================
import sys
from pathlib import Path


def main():
    if len(sys.argv) < 3:
        print("Usage: pdf_worker.py <src_drw> <dest_pdf>", file=sys.stderr)
        sys.exit(1)

    src  = Path(sys.argv[1])
    dest = Path(sys.argv[2])

    if not src.exists():
        print(f"File sorgente non trovato: {src}", file=sys.stderr)
        sys.exit(1)

    dest.parent.mkdir(parents=True, exist_ok=True)

    try:
        import win32com.client
        import pythoncom
    except ImportError:
        print("win32com non disponibile", file=sys.stderr)
        sys.exit(1)

    pythoncom.CoInitialize()
    sw   = None
    doc  = None
    src_str  = str(src).replace("/", "\\")
    dest_str = str(dest).replace("/", "\\")

    try:
        # ── Connessione a SolidWorks ────────────────────────────────────
        try:
            sw = win32com.client.GetActiveObject("SldWorks.Application")
            print("Connesso a SolidWorks esistente")
        except Exception:
            sw = win32com.client.Dispatch("SldWorks.Application")
            print("Avviato nuovo SolidWorks")

        sw.Visible = True

        # ── Apertura DRW ────────────────────────────────────────────────
        # OpenDoc6: Type=3 (swDocDRAWING), Options=1 (swOpenDocOptions_Silent)
        errors_v   = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
        warnings_v = win32com.client.VARIANT(
            pythoncom.VT_BYREF | pythoncom.VT_I4, 0)

        doc = sw.OpenDoc6(src_str, 3, 1, "", errors_v, warnings_v)
        print(f"OpenDoc6 errors={errors_v.value} warnings={warnings_v.value}")

        if doc is None:
            # Fallback: OpenDoc semplice (apre il documento già presente)
            doc = sw.OpenDoc(src_str, 3)
            print("Usato fallback OpenDoc")

        if doc is None:
            print(f"Impossibile aprire il DRW: {src_str}", file=sys.stderr)
            sys.exit(1)

        # Attiva il documento (necessario per export)
        try:
            err_act = win32com.client.VARIANT(
                pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            sw.ActivateDoc2(src_str, False, err_act)
        except Exception as ea:
            print(f"ActivateDoc2 non critico: {ea}")

        # Usa ActiveDoc dopo l'attivazione
        active = sw.ActiveDoc
        if active is not None:
            doc = active

        # ── Export PDF ─────────────────────────────────────────────────
        # Metodo 1: Extension.SaveAs (più moderno, raccomandato SW 2007+)
        ok = False
        try:
            export_data = sw.GetExportFileData(1)   # 1 = swExportPdfData
            if export_data is not None:
                try:
                    export_data.ViewPdfAfterSaving = False
                except Exception:
                    pass
            err2  = win32com.client.VARIANT(
                pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            warn2 = win32com.client.VARIANT(
                pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
            result = doc.Extension.SaveAs(
                dest_str,
                0,            # swSaveAsCurrentVersion
                1,            # swSaveAsOptions_Silent
                export_data,
                err2,
                warn2,
            )
            print(f"Extension.SaveAs result={result} err={err2.value} warn={warn2.value}")
            ok = bool(result)
        except Exception as e1:
            print(f"Extension.SaveAs fallito: {e1}")

        # Metodo 2: SaveAs3 senza export data (più compatibile)
        if not ok and not dest.exists():
            try:
                err3  = win32com.client.VARIANT(
                    pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                warn3 = win32com.client.VARIANT(
                    pythoncom.VT_BYREF | pythoncom.VT_I4, 0)
                result3 = doc.SaveAs3(dest_str, 0, 1, None, err3, warn3)
                print(f"SaveAs3 result={result3} err={err3.value} warn={warn3.value}")
                ok = bool(result3)
            except Exception as e2:
                print(f"SaveAs3 fallito: {e2}")

        # Metodo 3: SaveAs base (SolidWorks rileva .pdf dall'estensione)
        if not ok and not dest.exists():
            try:
                result4 = doc.SaveAs(dest_str)
                print(f"SaveAs result={result4}")
                ok = bool(result4)
            except Exception as e3:
                print(f"SaveAs fallito: {e3}")

        # ── Verifica risultato ──────────────────────────────────────────
        if dest.exists() and dest.stat().st_size > 0:
            print(f"OK: {dest_str}")
            sys.exit(0)
        else:
            print(
                f"PDF non creato (ok={ok}, exists={dest.exists()})",
                file=sys.stderr,
            )
            sys.exit(1)

    except Exception as e:
        print(f"Errore COM: {e}", file=sys.stderr)
        sys.exit(1)

    finally:
        if doc is not None and sw is not None:
            try:
                sw.CloseDoc(src_str)
                print("Documento chiuso")
            except Exception as ce:
                print(f"CloseDoc: {ce}")
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


if __name__ == "__main__":
    main()
