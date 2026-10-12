"""표준품셈 전문 목차(2026 전기공사)와 전기품셈_원표.json 의 품셈근거를 대조해 빠진 항목을 찾는다.

사용: python scripts/check_pumsam_toc.py data/전기품셈_원표.json
목차는 https://sec.kea.kr/sub_data/all.php?pcode=L 의 목차 PDF 기준입니다.
"""
import json,re,sys
def rng(ch,a,b,subs=None):
    subs=subs or {}
    out=[]
    for n in range(a,b+1):
        out.append(f"{ch}-{n}")
        for s in subs.get(n,[]): out.append(f"{ch}-{n}-{s}")
    return out
S=lambda n:list(range(1,n+1))
toc=[]
# 2장
toc+=rng(2,1,37,{1:S(3),2:S(3),4:S(3),5:S(4),6:S(7),7:[1],8:S(6),10:S(2),11:S(8),15:S(2),16:S(5),20:S(6),21:S(2),23:S(2),27:S(2),28:S(9),29:S(4),30:S(7),31:S(16),32:S(2),33:S(8),34:S(4)})
# 3장
sub3={4:[1],6:[1],8:[1],14:[1],28:[1,2],29:[1,2],39:[1],41:[1],42:[1],43:[1],44:[1],45:[1],46:[1],47:[1],50:[1],58:S(3),62:S(3),96:[1],131:[1],143:[1],152:[1,2],163:[1]}
toc+=rng(3,1,171,sub3)
# 4장
sub4={2:S(3),3:[1],5:[1,2],10:[1],20:[1],25:S(3),27:[1],30:[1],35:[1],36:[1],42:S(3),43:S(7),45:S(2),62:[1],68:S(7),76:[1],83:S(2),84:S(9),85:S(6),86:S(17),87:S(3),88:S(2),89:S(9),90:S(4),91:S(4),92:S(7),93:S(6),97:S(24)}
toc+=rng(4,1,97,sub4)
# 5장
sub5={3:[1],8:[1,2],9:[1],11:[1],18:S(3),21:S(4),25:S(3),26:S(8),27:S(7),29:[1,2],33:[1],42:[1,2],45:[1,2],46:S(4),47:[1,2],55:[1],21:S(4)}
toc+=rng(5,1,59,sub5)
toc+=rng(6,1,8)+rng(7,1,76,{16:[1],24:[1],31:[1],36:[1],58:[1],60:S(3)})+rng(8,1,9)+rng(9,1,3,{1:[1,2],3:[1]})
toc+=rng(10,1,4,{1:S(3),2:S(14),3:S(4),4:S(3)})
toc=set("전기"+t for t in toc)
d=json.load(open(sys.argv[1]))
db=set(r['품셈근거'].replace(' ','') for r in d)
print('목차',len(toc),'DB',len(db))
miss=sorted(toc-db,key=lambda k:[int(x) for x in re.findall(r'\d+',k)])
extra=sorted(db-toc,key=lambda k:[int(x) for x in re.findall(r'\d+',k)])
import collections
by=collections.defaultdict(list)
for m in miss: by[int(re.findall(r'\d+',m)[0])].append(m[2:])
for k,v in sorted(by.items()): print(f'{k}장 누락 {len(v)}:',' '.join(v))
print('DB에만 있음',[e[2:] for e in extra])
# 품이 비어있는 ref (항목은 있으나 수치 없음)
cnt=collections.Counter(); has=collections.Counter()
for r in d:
    k=r['품셈근거'].replace(' ','');cnt[k]+=1
    if isinstance(r.get('품셈'),(int,float)): has[k]+=1
nov=[k[2:] for k in cnt if has[k]==0]
print('수치 0건인 ref',len(nov),nov[:60])
