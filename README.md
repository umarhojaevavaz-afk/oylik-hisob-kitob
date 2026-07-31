# Moliya boti

Kunlik kirim/chiqimni Telegram orqali yozib, har oy uchun alohida Excel fayl yuritadigan bot.

## Ishga tushirish

Token endi kod ichida emas, muhit o'zgaruvchisidan olinadi:

```bash
export BOT_TOKEN="botfather_bergan_token"
pip install -r requirements.txt
python bot.py
```

Hosting (Railway/Heroku va h.k.) ishlatilsa, `BOT_TOKEN` ni panel ichidagi
**Environment Variables** bo'limiga qo'shing.

| O'zgaruvchi | Majburiy | Izoh |
|---|---|---|
| `BOT_TOKEN` | ha | @BotFather bergan token |
| `EXCEL_DIR` | yo'q | Excel fayllar papkasi (default: `moliya_data`) |

## Excel formati

Har oy uchun `moliya_data/YYYY-MM-moliya.xlsx` fayli yaratiladi:

| Sana | Vaqt | Tur | Kategoriya | Summa | Valyuta | Izoh |
|---|---|---|---|---|---|---|

Summa doim raqam bo'lib saqlanadi (valyuta alohida ustunda), shuning uchun
Excelda yig'indi, filtr va diagramma to'g'ri ishlaydi. Eski formatdagi fayllar
ochilganda sarlavhalari avtomatik yangilanadi, mavjud yozuvlarga tegilmaydi.

## Buyruqlar

- `/start` — asosiy menyu
- `/reset` — joriy oy ma'lumotlarini tozalash (tasdiq so'raydi, zaxira nusxa qoldiradi)
- `/cancel` — joriy amalni bekor qilish
