// Haberimiz Olsun izlenme Reels'i onay sayfası — yanıt kaydı (Google Apps Script web uygulaması).
// Kurulum: script.google.com → Yeni proje → bu kodu yapıştır → Dağıt → Yeni dağıtım → Web uygulaması
//   Yürüten: Ben · Erişim: Herkes → Dağıt → Web uygulaması URL'sini index.html içindeki ONAY_SCRIPT_URL'ye yaz
//   (ya da: node scripts/onay-site.mjs <URL>).
// İlk kayıtta e-tablo Drive'da KLASOR_YOLU altına oluşturulur (eksik klasörler açılır); kök klasöre düşmez.
// Her kayıt bir satır. İzin ekranında Drive ve E-Tablolar erişimini onaylayın.
const TABLO_ADI = 'Haberimiz Olsun - İzlenme Reels Onay Yanıtları';
const KLASOR_YOLU = ['Claude Code - Projects', 'YGTK', 'Haberimiz_Olsun_Project', 'Onay_Yanitlari'];
function klasor() {
  let k = DriveApp.getRootFolder();
  KLASOR_YOLU.forEach(ad => { const it = k.getFoldersByName(ad); k = it.hasNext() ? it.next() : k.createFolder(ad); });
  return k;
}
function tablo() {
  const p = PropertiesService.getScriptProperties();
  let id = p.getProperty('tabloId');
  let ss = null;
  try { ss = id ? SpreadsheetApp.openById(id) : null; } catch (e) { ss = null; }
  if (!ss) {
    ss = SpreadsheetApp.create(TABLO_ADI);
    DriveApp.getFileById(ss.getId()).moveTo(klasor());
    p.setProperty('tabloId', ss.getId());
    ss.getActiveSheet().appendRow(['zaman', 'no', 'ad', 'karar', 'not']);
  }
  return ss.getActiveSheet();
}
function doPost(e) {
  const veri = JSON.parse(e.postData.contents);
  const s = tablo();
  const zaman = new Date().toISOString();
  (veri.kayitlar || []).forEach(k => s.appendRow([zaman, String(k.no), String(k.ad || '').slice(0, 80), String(k.karar || ''), String(k.not || '').slice(0, 2000)]));
  return ContentService.createTextOutput(JSON.stringify({tamam: true})).setMimeType(ContentService.MimeType.JSON);
}
function doGet() {
  const s = tablo();
  const v = s.getDataRange().getValues().slice(1);
  const satirlar = v.map(r => ({zaman: r[0] instanceof Date ? r[0].toISOString() : String(r[0]), no: String(r[1]), ad: String(r[2]), karar: String(r[3]), not: String(r[4])}));
  return ContentService.createTextOutput(JSON.stringify({satirlar})).setMimeType(ContentService.MimeType.JSON);
}
// Kurulumdan sonra bir kez çalıştırılabilir: tabloyu hemen oluşturur ve klasörüne koyar.
function kur() { tablo(); }
