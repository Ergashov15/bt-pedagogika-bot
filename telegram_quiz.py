import os
import zipfile
import xml.etree.ElementTree as ET
import random
import time
import urllib.request
import urllib.parse
import json

# ===================== SOZLAMALAR =====================
BOT_TOKEN      = os.environ.get("BOT_TOKEN", "")
CHAT_ID        = os.environ.get("CHAT_ID", "")

START_FROM     = 1    # Qaysi savoldan boshlash
END_AT         = 160  # Qaysi savolda to'xtatish
ANSWER_TIMEOUT = 300  # Javob kutish vaqti (soniya)
# ======================================================

XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'BT PEDAGOGIKASIDAN TESTLAR!!!.xlsx')
_data_dir  = '/data' if os.path.isdir('/data') else os.path.dirname(os.path.abspath(__file__))
STATE_FILE = os.path.join(_data_dir, 'state.json')


def load_questions():
    with zipfile.ZipFile(XLSX) as z:
        shared_strings = []
        with z.open('xl/sharedStrings.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            for si in root.findall('ns:si', ns):
                texts = si.findall('.//ns:t', ns)
                shared_strings.append(''.join(t.text or '' for t in texts))

        with z.open('xl/worksheets/sheet1.xml') as f:
            tree = ET.parse(f)
            root = tree.getroot()
            ns = {'ns': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            rows = []
            for row in root.findall('.//ns:row', ns):
                row_data = []
                for cell in row.findall('ns:c', ns):
                    t = cell.get('t')
                    v_el = cell.find('ns:v', ns)
                    if v_el is not None:
                        val = v_el.text
                        if t == 's':
                            val = shared_strings[int(val)]
                    else:
                        val = ''
                    row_data.append(val)
                rows.append(row_data)

    questions = []
    for i, row in enumerate(rows):
        if i == 0:
            continue
        if i == 1:
            q       = row[1].strip() if len(row) > 1 else ''
            correct = row[2].strip() if len(row) > 2 else ''
            wrongs  = [row[j].strip() for j in range(3, min(6, len(row)))]
        else:
            q       = row[0].strip() if len(row) > 0 else ''
            correct = row[1].strip() if len(row) > 1 else ''
            wrongs  = [row[j].strip() for j in range(2, min(5, len(row)))]

        if not q or not correct:
            continue

        wrongs = [w for w in wrongs if w][:3]

        options = [correct] + wrongs
        random.seed(i * 17 + 3)
        random.shuffle(options)
        correct_idx = options.index(correct)

        q = q.replace('\n', ' ').strip()
        fixed_options = []
        for opt in options:
            opt = opt.replace('\n', ' ').strip()
            if len(opt) > 100:
                opt = opt[:97] + '...'
            fixed_options.append(opt)

        questions.append({
            'q':       q,
            'options': fixed_options,
            'ans':     correct_idx,
        })

    return questions


def load_state():
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE) as f:
                return json.load(f)
        except Exception:
            pass
    return {'next_q': START_FROM, 'offset': None, 'poll_id': None}


def save_state(next_q, offset, poll_id=None):
    with open(STATE_FILE, 'w') as f:
        json.dump({'next_q': next_q, 'offset': offset, 'poll_id': poll_id}, f)


def get_updates(offset=None, timeout=30):
    url = f'https://api.telegram.org/bot{BOT_TOKEN}/getUpdates'
    params = {'timeout': timeout, 'allowed_updates': json.dumps(['poll_answer'])}
    if offset is not None:
        params['offset'] = offset
    query = urllib.parse.urlencode(params)
    req = urllib.request.Request(f'{url}?{query}')
    try:
        with urllib.request.urlopen(req, timeout=timeout + 10) as resp:
            result = json.loads(resp.read().decode())
            return result.get('result', [])
    except Exception as e:
        print(f"  getUpdates xatosi: {e}")
        return []


def send_quiz(num, question, options, correct_option_id):
    url = f'https://api.telegram.org/bot{BOT_TOKEN}/sendPoll'

    q_with_num = f"{num}. {question}"
    if len(q_with_num) > 300:
        q_with_num = q_with_num[:297] + '...'

    data = {
        'chat_id':           CHAT_ID,
        'question':          q_with_num,
        'options':           json.dumps(options),
        'type':              'quiz',
        'correct_option_id': correct_option_id,
        'is_anonymous':      False,
    }

    encoded = urllib.parse.urlencode(data).encode('utf-8')
    req = urllib.request.Request(url, data=encoded)
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            result = json.loads(resp.read().decode())
            return result.get('ok', False), result
    except Exception as e:
        return False, str(e)


def wait_for_answer(poll_id, offset, timeout_sec=300):
    """Poll_id ga javob kelgunicha yoki timeout tugagunicha kutadi."""
    start = time.time()
    current_offset = offset
    while True:
        elapsed = time.time() - start
        if elapsed >= timeout_sec:
            break
        remaining = timeout_sec - elapsed
        wait = min(28, int(remaining))
        if wait <= 0:
            break
        updates = get_updates(offset=current_offset, timeout=wait)
        for update in updates:
            current_offset = update['update_id'] + 1
            if 'poll_answer' in update:
                if update['poll_answer']['poll_id'] == poll_id:
                    return current_offset
    return current_offset


def main():
    print("Savollar yuklanmoqda...")
    questions = load_questions()
    print(f"Jami: {len(questions)} ta savol")

    while True:
        state     = load_state()
        current_q = state.get('next_q', START_FROM)
        offset    = state.get('offset')
        poll_id   = state.get('poll_id')

        # Barcha savollar tugasa qaytadan boshlash
        if current_q > END_AT or current_q > len(questions):
            print("\nBarcha savollar yakunlandi, qaytadan boshlanmoqda...")
            save_state(START_FROM, offset, None)
            time.sleep(2)
            continue

        # Restart bo'lgan holat: oldingi poll hali javob kutayapti
        # Yangi savol yubormasdan, avval shu pollni yakunlaymiz
        if poll_id:
            print(f"[{current_q:>3}/{END_AT}] Oldingi savol — javob kutilmoqda...")
            offset = wait_for_answer(poll_id, offset, timeout_sec=ANSWER_TIMEOUT)
            save_state(current_q + 1, offset, None)
            continue

        # Birinchi marta ishga tushganda eski updatelarni tozalash
        if offset is None:
            old_updates = get_updates(offset=None, timeout=0)
            if old_updates:
                offset = old_updates[-1]['update_id'] + 1

        item = questions[current_q - 1]
        ok, resp = send_quiz(current_q, item['q'], item['options'], item['ans'])

        if ok:
            new_poll_id = resp['result']['poll']['id']
            # Poll yuborildi — state'ga saqlaymiz (restart bo'lsa ham ikkinchi savol chiqmasin)
            save_state(current_q, offset, new_poll_id)
            print(f"[{current_q:>3}/{END_AT}] Yuborildi — javob kutilmoqda...")
            offset = wait_for_answer(new_poll_id, offset, timeout_sec=ANSWER_TIMEOUT)
            save_state(current_q + 1, offset, None)
            print(f"[{current_q:>3}/{END_AT}] OK, keyingisiga o'tilmoqda")
        else:
            print(f"[{current_q:>3}/{END_AT}] XATO: {resp}")
            if 'Too Many Requests' in str(resp):
                print("  Rate limit — 30 soniya kutilmoqda...")
                time.sleep(30)
            else:
                time.sleep(5)


if __name__ == '__main__':
    main()
