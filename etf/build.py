# 把 etf/engine.js、webpush.js、gas.js 合併進 apps-script/Code.gs 最後面（Apps Script 只貼一個檔）。
# 改了 etf/ 裡的檔案後執行：python etf/build.py
import io, os, re
root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BEGIN = '// ===== 0050 長期投資（由 etf/build.py 從 etf/*.js 合併，要改請改 etf/ 裡的檔案）=====\n'
END = '// ===== 0050 長期投資 結束 =====\n'
parts = []
for name in ['engine.js', 'webpush.js', 'gas.js']:
    s = io.open(os.path.join(root, 'etf', name), encoding='utf-8').read()
    s = s.replace("'use strict';\n", '')
    s = re.sub(r"\nif \(typeof module !== 'undefined'\).*\n", '\n', s)
    parts.append('// ---- etf/' + name + ' ----\n' + s.strip() + '\n')
block = BEGIN + '\n'.join(parts) + END
p = os.path.join(root, 'apps-script', 'Code.gs')
code = io.open(p, encoding='utf-8', newline='').read()
if BEGIN in code:
    a = code.index(BEGIN); b = code.index(END) + len(END)
    code = code[:a] + block + code[b:]
else:
    code = code.rstrip('\n') + '\n\n' + block
io.open(p, 'w', encoding='utf-8', newline='').write(code)
print('ok', len(block))
