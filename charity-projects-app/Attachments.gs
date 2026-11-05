/**
 * المرحلة الرابعة: مرفقات السندات على Google Drive.
 * الملفات تُحفظ في Drive الخاص بحساب ناشر التطبيق داخل مجلد:
 *   مرفقات نظام المشروعات / <اسم الصفحة> / <رقم السند>
 * ولا تُشارك مع أحد؛ المستخدمون يفتحونها من خلال التطبيق وفق صلاحياتهم.
 */

function attachRoot_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('ATTACH_FOLDER_ID');
  if (id) {
    try { var f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) { /* أُنشئ من جديد */ }
  }
  var root = DriveApp.createFolder('مرفقات نظام المشروعات');
  props.setProperty('ATTACH_FOLDER_ID', root.getId());
  return root;
}

function subFolder_(parent, name) {
  var it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function docFolder_(entity, key) {
  return subFolder_(subFolder_(attachRoot_(), ENTITIES[entity].plural), String(key));
}

function listAttachments_(entity, key) {
  return readTable_(ATTACH.sheet)
    .filter(function (a) { return a.Entity === entity && String(a.DocKey) === String(key); })
    .map(function (a) {
      return {
        id: a.AttachmentId, name: a.FileName, mime: a.MimeType, size: Number(a.Size) || 0,
        at: toClient_(a.UploadedAt, 'datetime'), by: a.UploadedBy,
        url: 'https://drive.google.com/file/d/' + a.FileId + '/view'
      };
    });
}

function findAttachment_(id) {
  return readTable_(ATTACH.sheet).filter(function (a) { return String(a.AttachmentId) === String(id); })[0] || null;
}

/** صلاحية الرفع/الحذف: إضافة أو تعديل على صفحة السند */
function requireAttachEdit_(u, cfg) {
  if (!can_(u, cfg.page, 'add') && !can_(u, cfg.page, 'edit')) requirePerm_(u, cfg.page, 'edit');
}

/**
 * file = { name, mimeType, data(base64) }
 * يمكن إرفاق ملفات حتى بعد الترحيل (لا تغيّر القيم المالية)، أما الحذف من سند مرحّل فلمدير النظام فقط.
 */
function apiUploadAttachment(token, entity, key, file) {
  return run_(function () {
    var u = requireUser_(token);
    var cfg = docCfg_(entity);
    requireAttachEdit_(u, cfg);
    if (!findDoc_(cfg, key)) throw new Error('احفظ السند أولاً قبل إرفاق الملفات.');
    file = file || {};
    var name = String(file.name || '').replace(/[\\\/:*?"<>|]/g, '_').trim().substring(0, 150);
    var m = name.match(/\.([a-z0-9]+)$/i);
    var ext = m ? m[1].toLowerCase() : '';
    var mime = ATTACH.ext[ext];
    if (!name || !mime) {
      throw new Error('نوع الملف غير مسموح. المسموح: ' + Object.keys(ATTACH.ext).join(', '));
    }
    var bytes;
    try { bytes = Utilities.base64Decode(String(file.data || '')); } catch (e) { throw new Error('تعذر قراءة الملف.'); }
    if (!bytes.length) throw new Error('الملف فارغ.');
    if (bytes.length > ATTACH.maxMB * 1024 * 1024) throw new Error('حجم الملف أكبر من ' + ATTACH.maxMB + ' ميجابايت.');

    var created = docFolder_(entity, key).createFile(Utilities.newBlob(bytes, mime, name));
    created.setDescription(cfg.singular + ' ' + key + ' — رفع بواسطة ' + u.Username);
    return withLock_(function () {
      var rows = readTable_(ATTACH.sheet);
      appendObj_(ATTACH.sheet, {
        AttachmentId: nextId_(rows, 'AttachmentId'), Entity: entity, DocKey: String(key), FileId: created.getId(),
        FileName: name, MimeType: mime, Size: bytes.length, UploadedAt: new Date(), UploadedBy: u.Username
      });
      audit_(u, 'إرفاق ملف', entity, key, name);
      return listAttachments_(entity, key);
    });
  });
}

/** تنزيل الملف عبر التطبيق (يعمل لكل من لديه صلاحية العرض دون مشاركة الملف في Drive) */
function apiGetAttachment(token, id) {
  return run_(function () {
    var u = requireUser_(token);
    var a = findAttachment_(id);
    if (!a) throw new Error('المرفق غير موجود.');
    var cfg = docCfg_(a.Entity);
    requirePerm_(u, cfg.page, 'view');
    var blob;
    try { blob = DriveApp.getFileById(a.FileId).getBlob(); } catch (e) { throw new Error('الملف غير موجود في Drive، ربما حُذف يدوياً.'); }
    return { name: a.FileName, mime: a.MimeType, data: Utilities.base64Encode(blob.getBytes()) };
  });
}

function apiDeleteAttachment(token, id) {
  return run_(function () {
    var u = requireUser_(token);
    var a = findAttachment_(id);
    if (!a) throw new Error('المرفق غير موجود.');
    var cfg = docCfg_(a.Entity);
    requireAttachEdit_(u, cfg);
    var doc = findDoc_(cfg, a.DocKey);
    if (doc && doc.Status === 'Posted' && !isTrue_(u.IsAdmin)) throw new Error('حذف مرفقات السند المرحّل متاح لمدير النظام فقط.');
    return withLock_(function () {
      try { DriveApp.getFileById(a.FileId).setTrashed(true); } catch (e) { /* الملف محذوف مسبقاً */ }
      var cur = findAttachment_(id);
      if (cur) deleteRows_(ATTACH.sheet, [cur._row]);
      audit_(u, 'حذف مرفق', a.Entity, a.DocKey, a.FileName);
      return listAttachments_(a.Entity, a.DocKey);
    });
  });
}

/** يُستدعى عند حذف السندات: نقل ملفاتها لسلة المهملات وحذف سجلاتها */
function deleteDocAttachments_(entity, keys) {
  var rows = readTable_(ATTACH.sheet).filter(function (a) { return a.Entity === entity && keys.indexOf(String(a.DocKey)) >= 0; });
  rows.forEach(function (a) { try { DriveApp.getFileById(a.FileId).setTrashed(true); } catch (e) { /* تجاهل */ } });
  deleteRows_(ATTACH.sheet, rows.map(function (a) { return a._row; }));
}
