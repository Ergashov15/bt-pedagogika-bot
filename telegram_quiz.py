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
ANSWER_TIMEOUT = 300  # Javob kutish vaqti (soniya), keyin keyingisiga o'tadi
# ======================================================

XLSX = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'BT PEDAGOGIKASIDAN TESTLAR!!!.xlsx')


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
    start = time.time()
    current_offset = offset
    while time.time() - start < timeout_sec:
        remaining = timeout_sec - (time.time() - start)
        wait = min(30, int(remaining))
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

    batch = questions[START_FROM - 1: END_AT]
    print(f"{START_FROM} dan {START_FROM + len(batch) - 1} gacha {len(batch)} ta savol yuboriladi\n")

    # Eski updatelarni tozalash
    old_updates = get_updates(offset=None, timeout=0)
    offset = None
    if old_updates:
        offset = old_updates[-1]['update_id'] + 1

    ok_count = 0
    for i, item in enumerate(batch):
        num = START_FROM + i
        ok, resp = send_quiz(num, item['q'], item['options'], item['ans'])

        if ok:
            ok_count += 1
            poll_id = resp['result']['poll']['id']
            print(f"[{num:>3}/{END_AT}] Yuborildi — javob kutilmoqda...")
            offset = wait_for_answer(poll_id, offset, timeout_sec=ANSWER_TIMEOUT)
            print(f"[{num:>3}/{END_AT}] OK, keyingisiga o'tilmoqda")
        else:
            print(f"[{num:>3}/{END_AT}] XATO: {resp}")
            if 'Too Many Requests' in str(resp):
                print("  Rate limit — 30 soniya kutilmoqda...")
                time.sleep(30)
                ok2, resp2 = send_quiz(num, item['q'], item['options'], item['ans'])
                if ok2:
                    ok_count += 1
                    poll_id = resp2['result']['poll']['id']
                    offset = wait_for_answer(poll_id, offset, timeout_sec=ANSWER_TIMEOUT)

    print(f"\nYakunlandi: {ok_count}/{len(batch)} ta savol yuborildi.")


if __name__ == '__main__':
    main()
