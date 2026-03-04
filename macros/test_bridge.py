"""Test temporaneo per debug macro bridge."""
import sys, os, traceback
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from macros.sw_bridge import load_session, find_document

print("=== Test sw_bridge ===")

# 1. load_session
try:
    db, sp, user = load_session()
    print(f"[OK] DB connesso, User: {user['full_name']} (id={user['id']})")
except Exception as e:
    print(f"[FAIL] load_session: {e}")
    traceback.print_exc()
    sys.exit(1)

# 2. Elenca i documenti in DB
docs = db.fetchall("SELECT id, code, revision, doc_type, file_name, state, is_locked FROM documents LIMIT 10")
print(f"\n[INFO] {len(docs)} documenti in DB (primi 10):")
for d in docs:
    print(f"  id={d['id']} code={d['code']} rev={d['revision']} type={d['doc_type']} "
          f"file={d.get('file_name','')} state={d['state']} locked={d['is_locked']}")

# 3. Prova find_document con un file di test
if docs:
    test_doc = docs[0]
    from core.file_manager import EXT_FOR_TYPE
    ext = EXT_FOR_TYPE.get(test_doc['doc_type'], '.SLDPRT')
    fake_path = f"C:\\test\\{test_doc['code']}{ext}"
    print(f"\n[TEST] find_document('{fake_path}')")
    found = find_document(db, fake_path)
    if found:
        print(f"  [OK] Trovato: id={found['id']} code={found['code']}")
    else:
        print(f"  [FAIL] Non trovato")

# 4. Prova checkout (dry run - solo verifica)
if docs:
    test_doc = docs[0]
    print(f"\n[TEST] Checkout doc id={test_doc['id']} code={test_doc['code']} "
          f"state={test_doc['state']} locked={test_doc['is_locked']}")
    from core.checkout_manager import CheckoutManager, READONLY_STATES
    co = CheckoutManager(db, sp, user)
    if test_doc['state'] in READONLY_STATES:
        print(f"  [INFO] Stato '{test_doc['state']}' -> checkout bloccato (OK)")
    elif test_doc['is_locked']:
        print(f"  [INFO] Gia in checkout -> bloccato (OK)")
    else:
        print(f"  [INFO] Stato '{test_doc['state']}' -> checkout possibile")
        try:
            dest = co.checkout(test_doc['id'])
            print(f"  [OK] Checkout -> {dest}")
            # Undo immediato
            co.undo_checkout(test_doc['id'])
            print(f"  [OK] Undo checkout")
        except Exception as e:
            print(f"  [FAIL] {e}")
            traceback.print_exc()

print("\n=== Fine test ===")
