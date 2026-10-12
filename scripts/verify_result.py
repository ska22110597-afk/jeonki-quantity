"""결과 엑셀 정확성 검증.

사용: python scripts/verify_result.py 결과.xlsx 재계산본.xlsx 전기_표준품셈.xlsx
재계산본은 LibreOffice 로 변환해 수식값을 채운 파일입니다.
  soffice --headless --convert-to xlsx --outdir calc 결과.xlsx
검사: 단가 참조, 부속품비 15%·잡재료비 2%·공구손료 3%, 제외 규칙, 노임, 품셈 원표, 합계, 목록 연결.
"""
import sys, re, collections, openpyxl
src, calc, db = sys.argv[1:4]
wf = openpyxl.load_workbook(src)            # 수식
wv = openpyxl.load_workbook(calc, data_only=True)  # 계산값
dbws = openpyxl.load_workbook(db, read_only=True)['품셈표']
issues = collections.defaultdict(list)
def bad(k, msg): issues[k].append(msg)

# --- 품셈 DB (명칭, 규격) -> [(노무, 품)]
pm = collections.defaultdict(list); last = [None, None]
for i, r in enumerate(dbws.iter_rows(min_row=2, values_only=True)):
    _, nm, sp, un, job, qty, rate, ref = (list(r) + [None] * 8)[:8]
    if nm: last[0] = nm
    if sp is not None or nm: last[1] = sp
    if job: pm[(last[0], last[1])].append((job, qty, rate))

wage = {r[0]: r[1] for r in wv['노임단가'].iter_rows(min_row=1, values_only=True) if r[0]}
cmp_f, cmp_v = wf['단가대비표'], wv['단가대비표']
il_f, il_v = wf['일위대가'], wv['일위대가']

blocks = []; cur = None
for r in range(5, il_f.max_row + 1):
    a = il_f.cell(r, 1).value
    if a and '호표' in str(a): cur = {'title': a, 'rows': []}; blocks.append(cur); continue
    if cur is None and a not in (None, ' '):
        cur = {'title': '(호표 1 제목행 없음)', 'rows': []}; blocks.append(cur); bad('호표 제목행 누락', f'{r}행 앞에 호표 제목 없음')
    if cur is not None and a not in (None, ' '): cur['rows'].append(r)
print('호표 수', len(blocks))

sys.path.insert(0, __import__('os').path.dirname(__import__('os').path.dirname(__import__('os').path.abspath(__file__))))
from app.fitting_rules import fitting_excluded
CONDUIT = lambda n: '전선관' in str(n)
WIRE = lambda n: ('전선' in str(n) or '케이블' in str(n)) and not CONDUIT(n) and not any(w in str(n) for w in ('트레이','덕트','몰딩'))
stats = collections.Counter()
for b in blocks:
    b['rows'] = [r for r in b['rows']]; sumrow = b['rows'].pop() if '합' in str(il_f.cell(b['rows'][-1],1).value) else None; rows = b['rows']; b['sum']=sumrow; m = rows[0]; name = il_f.cell(m, 1).value; spec = il_f.cell(m, 2).value
    tag = f"{b['title']}"
    names = [il_f.cell(r, 1).value for r in rows]
    # 1 자재 단가 참조
    ref = il_f.cell(m, 5).value
    mm = re.match(r"='?단가대비표'?!L(\d+)$", str(ref))
    if not mm: bad('단가참조', f'{tag} {ref}')
    else:
        cr = int(mm.group(1))
        if (cmp_f.cell(cr, 1).value, cmp_f.cell(cr, 2).value) != (name, spec): bad('단가참조', f'{tag} → 단가대비표 {cr}행 불일치')
        if il_v.cell(m, 5).value != cmp_v.cell(cr, 12).value: bad('단가값', tag)
    # 2 부속품비/잡재료비 유무·금액
    has_fit = '전선관부속품비' in names; has_sun = '잡재료비' in names
    if CONDUIT(name) and not has_fit and not fitting_excluded(name, spec): bad('부속품비 누락', tag)
    if has_fit and fitting_excluded(name, spec): bad('제외대상인데 부속품비 있음', tag)
    if has_fit and not CONDUIT(name): bad('부속품비 불필요(전선관 아님)', tag)
    if (CONDUIT(name) or WIRE(name)) and not has_sun: bad('잡재료비 누락', tag)
    if has_sun and not (CONDUIT(name) or WIRE(name)): bad('잡재료비 불필요', tag)
    mat_amt = il_v.cell(m, 6).value or 0
    for r in rows:
        n = il_f.cell(r, 1).value
        fr = 0.40 if ('CD' in f'{name} {spec}'.upper() or '주름' in str(name) or ('합성수지' in str(name) and '가요' in str(name))) else 0.15
        if n == '전선관부속품비' and abs((il_v.cell(r, 6).value or 0) - int(round(mat_amt * fr * 10, 6)) / 10) > 0.15: bad('부속품비 금액', tag)
        if n == '잡재료비' and abs((il_v.cell(r, 6).value or 0) - int(mat_amt * 0.02 * 10) / 10) > 0.05: bad('잡재료비 금액', tag)
    # 3 인부 품·노임
    labor = [r for r in rows if il_f.cell(r, 3).value == '인']
    key = (name, spec)
    exp = pm.get(key)
    got = [(il_f.cell(r, 1).value, il_v.cell(r, 4).value) for r in labor]
    if exp:
        stats['품셈 정확일치'] += 1
        if sorted((j, round(float(q or 0), 4)) for j, q, _ in exp) != sorted((j, round(float(q or 0), 4)) for j, q in got):
            bad('품 불일치', f'{tag} 기대{[(j,q) for j,q,_ in exp]} 실제{got}')
    else:
        stats['품셈 정확일치 없음(별칭/이하/기본값)'] += 1
        if len(got) == 1 and (got[0][1] or 0) == 0: bad('품 0 (품셈 없음 · 직접입력 필요)', tag)
    for r in labor:
        j = il_f.cell(r, 1).value
        if j not in wage: bad('노임표에 없는 직종', f'{tag} {j}')
        elif il_v.cell(r, 7).value != wage[j]: bad('노임 불일치', f'{tag} {j}')
    # 4 공구손료 3%
    tl = [r for r in rows if il_f.cell(r, 1).value == '공구손료']
    lab_amt = sum(il_v.cell(r, 8).value or 0 for r in labor)
    if not tl: bad('공구손료 누락', tag)
    elif abs((il_v.cell(tl[0], 10).value or 0) - int(lab_amt * 0.03 * 10) / 10) > 0.05: bad('공구손료 금액', tag)
    # 5 합계
    sr = sumrow
    tot = sum((il_v.cell(r, 12).value or 0) for r in rows)
    if abs((il_v.cell(sr, 12).value or 0) - tot) > 0.15: bad('합계', f'{tag} {il_v.cell(sr,12).value} vs {tot}')
# 6 일위대가목록 ↔ 호표 합계
lf, lv = wf['일위대가목록'], wv['일위대가목록']
sums = [b['sum'] for b in blocks]
for i, sr in enumerate(sums):
    r = 5 + i
    if lv.cell(r, 1).value is None: break
    if abs((lv.cell(r, 11).value or 0) - (il_v.cell(sr, 12).value or 0)) > 0.15 and (lv.cell(r,4).value in (None,0,1)):
        bad('목록≠호표합계', f'목록{r}행')
print(dict(stats))
for k, v in issues.items():
    print(f'\n[{k}] {len(v)}건'); [print('  ', x) for x in v[:12]]
if not issues: print('이상 없음')
