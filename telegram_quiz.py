import os
import zipfile
import xml.etree.ElementTree as ET
import random
import time
import urllib.request
import urllib.parse
import json

# ===================== SOZLAMALAR =====================
BOT_TOKEN = os.environ.get("BOT_TOKEN", "")
CHAT_ID   = os.environ.get("CHAT_ID", "")
                                         # Guruh: -xxxxxxxxxx
                                         # Shaxsiy: 123456789

START_FROM = 1     # Qaysi savoldan boshlash (1 = boshidan)
END_AT     = 160   # Qaysi savolda to'xtatish (160 = oxirigacha)
DELAY_SEC  = 3     # Har bir savol o'rtasidagi kutish (soniya)
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

        # Telegram: savol max 300 belgi, variant max 100 belgi
        q = q.replace('\n', ' ').strip()
        if len(q) > 300:
            q = q[:297] + '...'

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


def send_quiz(question, options, correct_option_id):
    url = f'https://api.telegram.org/bot{BOT_TOKEN}/sendPoll'

    data = {
        'chat_id':           CHAT_ID,
        'question':          question,
        'options':           json.dumps(options),
        'type':              'quiz',
        'correct_option_id': correct_option_id,
        'is_anonymous':      True,
    }

    encoded = urllib.parse.urlencode(data).encode('utf-8')
    req = urllib.request.Request(url, data=encoded)

    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            result = json.loads(resp.read().decode())
            return result.get('ok', False), result
    except Exception as e:
        return False, str(e)


def main():
    if BOT_TOKEN.startswith('SIZ') or CHAT_ID.startswith('SIZ'):
        print("XATO: BOT_TOKEN va CHAT_ID ni to'ldiring!")
        print("  BOT_TOKEN = @BotFather dan olingan token")
        print("  CHAT_ID   = guruh/kanal/foydalanuvchi ID")
        return

    print("Savollar yuklanmoqda...")
    questions = load_questions()
    print(f"Jami: {len(questions)} ta savol")

    batch = questions[START_FROM - 1: END_AT]
    print(f"{START_FROM} dan {START_FROM + len(batch) - 1} gacha {len(batch)} ta savol yuboriladi")
    print(f"Taxminiy vaqt: {len(batch) * DELAY_SEC // 60} daqiqa {len(batch) * DELAY_SEC % 60} soniya\n")

    ok_count = 0
    for i, item in enumerate(batch):
        num = START_FROM + i
        ok, resp = send_quiz(item['q'], item['options'], item['ans'])

        if ok:
            ok_count += 1
            print(f"[{num:>3}/{END_AT}] OK")
        else:
            print(f"[{num:>3}/{END_AT}] XATO: {resp}")
            # Rate limit bo'lsa biroz ko'proq kutamiz
            if 'Too Many Requests' in str(resp):
                print("  Rate limit - 30 soniya kutilmoqda...")
                time.sleep(30)
                # Qayta urinish
                ok2, resp2 = send_quiz(item['q'], item['options'], item['ans'])
                if ok2:
                    ok_count += 1
                    print(f"  Qayta urinish: OK")

        if i < len(batch) - 1:
            time.sleep(DELAY_SEC)

    print(f"\nYakunlandi: {ok_count}/{len(batch)} ta savol yuborildi.")


if __name__ == '__main__':
    main()
