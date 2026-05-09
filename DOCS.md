لكي تتمكن من استخدام سيرفر Facebook Ads MCP الذي قمنا بضبطه على أي جهاز كمبيوتر آخر، تحتاج فقط إلى نقل المكونات الأساسية الثلاثة: ملفات السيرفر، التوكن (Access Token)، وإعدادات Claude Desktop.

إليك الخطوات العملية لتشغيله على أي جهاز في العالم:

1. نقل ملفات السيرفر (الكود)
قم بنسخ مجلد fb-mcp-server الموجود لديك في d:\facebook_ads\fb-mcp-server. يمكنك رفعه على حسابك في GitHub (مستودع خاص Private)، أو وضعه على Google Drive/فلاش ميموري (USB). تأكد من الاحتفاظ بتعديلات ملف server.py التي قمنا بها (تحويل print إلى sys.stderr).

2. إعداد الجهاز الجديد
على أي جهاز كمبيوتر جديد (سواء Windows أو Mac أو Linux)، قم بالآتي:

قم بتحميل وتثبيت Python (نسخة 3.10 أو أحدث).
انسخ مجلد fb-mcp-server إلى الجهاز الجديد.
افتح موجه الأوامر (Terminal أو CMD) داخل هذا المجلد واكتب الأمر التالي لتثبيت المكتبات المطلوبة:
bash
pip install -r requirements.txt
3. إعداد Claude Desktop على الجهاز الجديد
قم بتثبيت تطبيق Claude Desktop.
افتح ملف إعدادات Claude على الجهاز الجديد. (تختلف مساراته حسب نظام التشغيل):
ويندوز (نسخة الموقع): %APPDATA%\Claude\claude_desktop_config.json
ويندوز (نسخة المتجر): LocalAppData\Packages\Claude_pzs8sxrjxfjjc\LocalCache\Roaming\Claude\claude_desktop_config.json
الماك (Mac): ~/Library/Application Support/Claude/claude_desktop_config.json
اللينكس: ~/.config/Claude/claude_desktop_config.json
أضف هذا الكود إلى الملف مع تغيير المسارات لتناسب الجهاز الجديد:
json
{
  "mcpServers": {
    "facebook-ads": {
      "command": "python",
      "args": [
        "مسار_الملف_على_الجهاز_الجديد/fb-mcp-server/server.py",
        "--fb-token",
        "التوكن_الخاص_بك_EAAb8rNYkc3ABRY0SQR..."
      ]
    }
  }
}
(ملاحظة: إذا كان الجهاز الجديد Mac أو Linux، استخدم python3 بدلاً من python في خانة command).

4. حماية التوكن (مهم جداً 🔒)
التوكن (EAAb...) الذي استخرجته قوي جداً ويملك صلاحيات كاملة لحسابك الإعلاني وإدارة البزنس على فيسبوك.

لا ترفعه أبداً على مستودع GitHub عام (Public).
إذا أردت تشغيل السيرفر على جهاز ليس ملكك، احذف التوكن من ملف الإعدادات بعد الانتهاء من العمل.
بهذه الخطوات الأربع، يصبح هذا الأداة (MCP) معك في أي مكان يعمل على أي جهاز لتدير حملاتك الإعلانية عبر Claude بكل سهولة!

