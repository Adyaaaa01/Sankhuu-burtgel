# Vibe Cafe Finance Web App

Vercel дээр оруулах demo app.

## Ашиглах
1. `npm install`
2. `npm run dev`
3. Browser дээр нээгээд `.xlsx` банкны хуулга upload хийнэ.
4. `Банк` sheet: Огноо, Гүйлгээний утга, Орлого, Зарлага, Код гэсэн дарааллаар уншина.

## Vercel deploy
- GitHub repository болгоод Vercel дээр Import хийнэ.
- Framework preset: Vite
- Build command: `npm run build`
- Output directory: `dist`

## Логик
- Орлого: Дт 102 Харилцах / Кт тухайн орлогын код
- Зарлага: Дт тухайн зарлагын код / Кт 102 Харилцах
- Ингэснээр Journal мөр бүр Дт=Кт болж тэнцэнэ.
- OZT, T данс, Balance, НӨАТ тулгалт, Dashboard автоматаар шинэчлэгдэнэ.
