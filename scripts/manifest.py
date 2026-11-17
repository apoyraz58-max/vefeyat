"""Generates manifest/manifest.template.xml (and shortcuts.json): the Vefeyât ribbon in the VBA 0.3.41 layout.

Run after changing the ribbon:  python3 scripts/manifest.py
Office add-in ribbons only allow buttons and menus (no drop-down boxes showing a value),
so the VBA setting boxes become menus; the pane shows the current values.
"""
from pathlib import Path

B = "{{BASE_URL}}"
short, long_, urls, used_icons = {}, {}, {}, []


def S(key, value):
    short[key] = value
    return key


def L(key, value):
    long_[key] = value
    return key


def U(key, path):
    urls[key] = B + path
    return key


def icon(name, ind):
    if name not in used_icons:
        used_icons.append(name)
    return f"{ind}<Icon>\n" + "".join(f'{ind}  <bt:Image size="{s}" resid="Ic.{name}.{s}"/>\n' for s in (16, 32, 80)) + f"{ind}</Icon>\n"


def head(label, tip, ind):
    return (f'{ind}<Label resid="{label}"/>\n{ind}<Supertip>\n{ind}  <Title resid="{label}"/>\n'
            f'{ind}  <Description resid="{tip}"/>\n{ind}</Supertip>\n')


def pane(url):
    return lambda ind: (f'{ind}<Action xsi:type="ShowTaskpane">\n{ind}  <TaskpaneId>VefeyatPane</TaskpaneId>\n'
                        f'{ind}  <SourceLocation resid="{url}"/>\n{ind}</Action>\n')


def function(name):
    return lambda ind: f'{ind}<Action xsi:type="ExecuteFunction">\n{ind}  <FunctionName>{name}</FunctionName>\n{ind}</Action>\n'


C, I = " " * 16, " " * 20


def button(cid, label, tip, ic, action, ind=C, kind="Button", tag="Control"):
    attr = f' xsi:type="{kind}"' if tag == "Control" else ""
    return f'{ind}<{tag}{attr} id="Vefeyat.{cid}">\n' + head(label, tip, ind + "  ") + icon(ic, ind + "  ") + action(ind + "  ") + f"{ind}</{tag}>\n"


def item(cid, label, tip, ic, action):
    return button(cid, label, tip, ic, action, I, tag="Item")


def menu(cid, label, tip, ic, items, ind=C):
    return (f'{ind}<Control xsi:type="Menu" id="Vefeyat.{cid}">\n' + head(label, tip, ind + "  ") + icon(ic, ind + "  ")
            + f"{ind}  <Items>\n" + "".join(items) + f"{ind}  </Items>\n{ind}</Control>\n")


def group(gid, label, ic, controls):
    ind = " " * 14
    return f'{ind}<Group id="Vefeyat.G.{gid}">\n{ind}  <Label resid="{S("G." + gid, label)}"/>\n' + icon(ic, ind + "  ") + "".join(controls) + f"{ind}</Group>\n"


def setting_menu(key, label, tip, ic, choices):
    tip_id = L(f"S.{key}.Tip", tip + " Seçili değer Vefeyât penceresinin üstünde yazar.")
    items = [item(f"S.{key}.{v}", S(f"S.{key}.{v}", text), tip_id, ic, function(f"set_{key}_{v}")) for v, text in choices]
    return menu(f"S.{key}", S(f"S.{key}", label), tip_id, ic, items)


insert = button("Insert", S("Insert", "Ölüm tarihi ekle"),
                L("Insert.Tip", "Seçili adın ölüm tarihini ekler. Tek ve kesin eşleşmede, ilk geçişte ve adın arkasında tarih yoksa doğrudan ekler; aksi hâlde Vefeyât penceresinde seçim veya onay ister."),
                "insertbrand", pane(U("U.Insert", "taskpane.html?action=insert")))
spelling = button("Spelling", S("Spelling", "İsim yazımı"),
                  L("Spelling.Tip", "Seçili adı kısa, tam ve orijinal adlarla karşılaştırır; Türkçe eki koruyarak yalnız seçili adı değiştirir."),
                  "spelling", pane(U("U.Spelling", "taskpane.html?action=spelling")))

groups = [
    group("Add", "Ekle ve denetle", "brand", [
        insert,
        button("Audit", S("Audit", "Tüm belgeyi denetle"), L("Audit.Tip", "Belgedeki adları sırayla gösterir; onaylayıp ilerlersiniz."), "audit", pane(U("U.Audit", "taskpane.html?view=denetim"))),
        spelling,
    ]),
    group("Review", "Denetim", "scope", [
        setting_menu("scope", "Tarih ekleme", "İlk kullanım (varsayılan), bütün kullanımlar veya yalnız denetim.", "scope",
                     [("first", "Yalnız ilk kullanım"), ("chapter", "Her bölümde ilk kullanım"), ("all", "Bütün kullanımlar"), ("none", "Yalnız denetle")]),
        setting_menu("auditMode", "Denetlenecek", "Tüm belgeyi denetle komutunun hangi geçişleri göstereceği.", "auditmode",
                     [("all", "Tüm geçişler"), ("undated", "Tarihi olmayanlar")]),
        button("Resume", S("Resume", "Denetime devam et"), L("Resume.Tip", "Duraklatılmış belge denetimini kaldığı geçişten sürdürür."), "resume", pane(U("U.Resume", "taskpane.html?view=devam"))),
        button("Bulk", S("Bulk", "Toplu düzeltme önizlemesi"), L("Bulk.Tip", "Yazım ve tarih önerilerini tek listede gösterir; yalnız seçtiklerinizi uygular."), "bulk", pane(U("U.Bulk", "taskpane.html?view=toplu"))),
    ]),
    group("Source", "Şahıs ve kaynak", "search", [
        button("Search", S("Search", "Şahıs ara"), L("Search.Tip", "Yerel katalogda ad arar; sonuçları DİA bağlantısıyla gösterir."), "search", pane(U("U.Search", "taskpane.html?view=ara"))),
        button("Dia", S("Dia", "DİA maddesini aç"), L("Dia.Tip", "Seçili şahsın DİA maddesinin bağlantısını gösterir."), "dia", pane(U("U.Dia", "taskpane.html?action=dia"))),
        button("SourceInfo", S("SourceInfo", "Tarihin kaynağı"), L("SourceInfo.Tip", "Seçili Vefeyât tarihinin hangi katalog kaydından eklendiğini ve güncel olup olmadığını gösterir."), "source", pane(U("U.SourceInfo", "taskpane.html?view=kaynak"))),
    ]),
    group("Document", "Belge", "inventory", [
        button("Inventory", S("Inventory", "Şahıs dökümü"), L("Inventory.Tip", "Metindeki şahıs adlarını ve tarihleri listeler."), "inventory", pane(U("U.Inventory", "taskpane.html?view=dokum"))),
        menu("More", S("More", "Diğer araçlar"), L("More.Tip", "Tarih ve katalog araçları, kimlik eşleşmeleri ve belge istisnaları."), "more", [
            item("Refresh", S("Refresh", "Tarihleri güncelle"), L("Refresh.Tip", "Tarih ve yazım önerilerini toplu önizlemede gösterir; mevcut tarihleri kendiliğinden yeniden yazmaz."), "more", pane(U("U.Refresh", "taskpane.html?view=toplu"))),
            item("Compare", S("Compare", "Katalogla karşılaştır"), L("Compare.Tip", "Belgedeki Vefeyât tarihlerini güncel katalogla karşılaştırır (Belge sağlığı)."), "more", pane(U("U.Compare", "taskpane.html?view=saglik"))),
            item("Registry", S("Registry", "Dinamik şahıs kayıtları"), L("Registry.Tip", "Geçişlerin bağlı olduğu katalog kimliklerini gösterir ve düzeltir."), "more", pane(U("U.Registry", "taskpane.html?view=kimlik"))),
            item("Coverage", S("Coverage", "Denetim kapsamını göster"), L("Coverage.Tip", "Taranan ve taranmayan belge bölümlerini raporlar."), "more", pane(U("U.Coverage", "taskpane.html?view=kapsam"))),
            item("Identity", S("Identity", "Kimlik eşleşmelerini düzelt"), L("Identity.Tip", "Belirsiz adlar için verdiğiniz kimlik onaylarını listeler; yanlış olanı silebilirsiniz."), "more", pane(U("U.Identity", "taskpane.html?view=kimlik"))),
            item("Exclusions", S("Exclusions", "Belge istisnaları"), L("Exclusions.Tip", "Şahıs değil olarak işaretlenen ifadeleri yönetir."), "more", pane(U("U.Exclusions", "taskpane.html?view=istisna"))),
        ]),
        button("Health", S("Health", "Belge sağlığı"), L("Health.Tip", "Vefeyât tarihlerini katalog ve önek ayarıyla karşılaştırır; yalnız seçtiklerinizi onarır."), "health", pane(U("U.Health", "taskpane.html?view=saglik"))),
        button("History", S("History", "İşlem geçmişi"), L("History.Tip", "Bu oturumdaki Vefeyât işlemlerini listeler; eklenen tarihlere gidebilirsiniz."), "history", pane(U("U.History", "taskpane.html?view=gecmis"))),
    ]),
    group("Submit", "Teslim ve dizin", "tools", [
        button("Consistency", S("Consistency", "Tutarlılık raporu"), L("Consistency.Tip", "ö./öl./v. ve tarih biçimlerinin karışık kullanıldığı yerleri gösterir; belgeyi değiştirmez."), "consistency", pane(U("U.Consistency", "taskpane.html?tool=tutarlilik"))),
        button("Report", S("Report", "Şahıs raporu"), L("Report.Tip", "Şahısları ilk geçiş sırasıyla ve kronolojik olarak yeni bir belgede listeler."), "report", pane(U("U.Report", "taskpane.html?tool=rapor"))),
        menu("Index", S("Index", "Şahıslar dizini"), L("Index.Tip", "Kimliği kesin geçişlerden Word şahıslar dizini."), "index", [
            item("IndexBuild", S("IndexBuild", "Dizini oluştur / güncelle"), "Index.Tip", "index", pane(U("U.IndexBuild", "taskpane.html?tool=dizin"))),
            item("IndexRemove", S("IndexRemove", "Dizin işaretlerini kaldır"), L("IndexRemove.Tip", "Gizli dizin işaretlerini siler; dizin, güncel hâliyle metin olarak kalır."), "index", pane(U("U.IndexRemove", "taskpane.html?tool=dizinkaldir"))),
            item("IndexDelete", S("IndexDelete", "Dizini tamamen kaldır"), L("IndexDelete.Tip", "Vefeyât dizinini, başlığını ve dizin işaretlerini birlikte siler."), "index", pane(U("U.IndexDelete", "taskpane.html?tool=dizinsil"))),
        ]),
        button("Finalize", S("Finalize", "Teslim öncesi temizle"), L("Finalize.Tip", "Teslim kopyası için Vefeyât izlerini kaldırır; görünen metin değişmez."), "finalize", pane(U("U.Finalize", "taskpane.html?tool=temizlik"))),
    ]),
    group("Settings", "Tercihler", "nameform", [
        setting_menu("nameForm", "Ad biçimi", "İsim yazımında önce önerilecek biçim.", "nameform",
                     [("ask", "Her seferinde sor"), ("near", "Seçili ada yakın"), ("short", "Kısa ad"), ("full", "Tam / uzun ad"), ("original", "Orijinal ad")]),
        setting_menu("prefix", "Kısaltma", "Ölüm tarihinin başındaki kısaltma.", "prefix",
                     [("original", "Kaynakta olduğu gibi"), ("o", "ö."), ("ol", "öl."), ("v", "v.")]),
        menu("Prefs", S("Prefs", "Ayarlar"), L("Prefs.Tip", "Yazım tercihleri ve görünüm ayarları."), "settings", [
            item("Prefer", S("Prefer", "Seçili yazımı tercih et"), L("Prefer.Tip", "Seçili adın yazımını bu şahıs için belgeye özgü tercih olarak kaydeder."), "settings", pane(U("U.Prefer", "taskpane.html?action=prefer"))),
            item("ClearSpelling", S("ClearSpelling", "Yazım tercihlerini temizle"), L("ClearSpelling.Tip", "Belgeye özgü yazım tercihlerini siler."), "settings", pane(U("U.ClearSpelling", "taskpane.html?action=clearSpelling"))),
            item("ExportPrefs", S("ExportPrefs", "Tercihleri dışa aktar"), L("ExportPrefs.Tip", "Ayarları, yazım tercihlerini, istisnaları ve kimlik onaylarını bir dosyaya kaydeder."), "settings", pane(U("U.ExportPrefs", "taskpane.html?view=tercihler"))),
            item("ImportPrefs", S("ImportPrefs", "Tercihleri al"), L("ImportPrefs.Tip", "Dışa aktarılmış tercih dosyasını bu belgeye uygular."), "settings", pane(U("U.ImportPrefs", "taskpane.html?view=tercihler"))),
            item("Shading", S("Shading", "Tarih alanlarını göster / gizle"), L("Shading.Tip", "Vefeyât tarihlerinin çevresine çerçeve koyar ya da kaldırır."), "settings", pane(U("U.Shading", "taskpane.html?action=shading"))),
            item("NoProof", S("NoProof", "Şahıs adlarını yazım denetiminden çıkar"), L("NoProof.Tip", "Bulunan şahıs adlarının kırmızı çizgisini kaldırır; sağ tıkta Vefeyât menüsü görünür."), "settings", pane(U("U.NoProof", "taskpane.html?action=noproof"))),
        ]),
    ]),
    group("Data", "Veri ve yardım", "info", [
        button("Info", S("Info", "Veri bilgisi"), L("Info.Tip", "Yüklü kataloğu ve sürümü gösterir."), "info", pane(U("U.Info", "taskpane.html?view=katalog"))),
        button("Csv", S("Csv", "CSV seç"), L("Csv.Tip", "Bilgisayarınızdaki başka bir DİA CSV dosyasını kullanır."), "csv", pane(U("U.Csv", "taskpane.html?view=katalog"))),
        button("Help", S("Help", "Kullanım kılavuzu"), L("Help.Tip", "Vefeyât komutlarının kısa kılavuzu ve Hakkında."), "brand", pane(U("U.Help", "taskpane.html?view=yardim"))),
    ]),
]

ribbon = ("          <ExtensionPoint xsi:type=\"PrimaryCommandSurface\">\n            <CustomTab id=\"Vefeyat.Tab\">\n"
          + "".join(groups) + f'              <Label resid="{S("Tab", "Vefeyât")}"/>\n            </CustomTab>\n          </ExtensionPoint>\n')
context = ("          <ExtensionPoint xsi:type=\"ContextMenu\">\n            <OfficeMenu id=\"ContextMenuText\">\n"
           + menu("Context", S("Context", "Vefeyât"), L("Context.Tip", "Seçili şahıs adı için Vefeyât komutları."), "brand", [
               item("C.Insert", "Insert", "Insert.Tip", "insertbrand", pane("U.Insert")),
               item("C.Spelling", "Spelling", "Spelling.Tip", "spelling", pane("U.Spelling")),
               item("C.Dia", "Dia", "Dia.Tip", "dia", pane("U.Dia")),
               item("C.Source", "SourceInfo", "SourceInfo.Tip", "source", pane("U.SourceInfo")),
           ], " " * 14)
           + "            </OfficeMenu>\n          </ExtensionPoint>\n")

U("Commands.Url", "commands.html")
short["GetStarted.Title"] = "Vefeyât hazır"
long_["GetStarted.Description"] = "Metinde bir şahıs adı seçin ve üstteki Vefeyât sekmesinden Ölüm tarihi ekle düğmesine basın."


def esc(value):
    return value.replace("&", "&amp;").replace('"', "&quot;").replace("<", "&lt;")


images = "".join(f'        <bt:Image id="Ic.{n}.{s}" DefaultValue="{B}assets/icons/{n}-{s}.png"/>\n' for n in used_icons for s in (16, 32, 80))
url_xml = "".join(f'        <bt:Url id="{k}" DefaultValue="{v}"/>\n' for k, v in urls.items())
url_xml += '        <bt:Url id="GetStarted.LearnMoreUrl" DefaultValue="https://github.com/apoyraz58-max/vefeyat"/>\n'
short_xml = "".join(f'        <bt:String id="{k}" DefaultValue="{esc(v)}"/>\n' for k, v in short.items())
long_xml = "".join(f'        <bt:String id="{k}" DefaultValue="{esc(v)}"/>\n' for k, v in long_.items())

xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<!-- Generated by scripts/manifest.py; rendered by scripts/build.mjs ({{{{BASE_URL}}}}, {{{{VERSION}}}}). -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
           xsi:type="TaskPaneApp">
  <Id>6f3c2a4e-8b1d-4c7a-9e52-0d9a7b3e5f21</Id>
  <Version>{{{{VERSION}}}}</Version>
  <ProviderName>Abdurrahman Poyraz</ProviderName>
  <DefaultLocale>tr-TR</DefaultLocale>
  <DisplayName DefaultValue="Vefeyât"/>
  <Description DefaultValue="Şahıs adlarını DİA kataloğuyla denetler, doğru yazımı önerir ve vefat tarihini ekler."/>
  <IconUrl DefaultValue="{B}assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{B}assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://github.com/apoyraz58-max/vefeyat"/>
  <AppDomains>
    <AppDomain>https://islamansiklopedisi.org.tr</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.3">
      <Set Name="WordApi" MinVersion="1.3"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="{B}taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Document">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="GetStarted.LearnMoreUrl"/>
          </GetStarted>
          <FunctionFile resid="Commands.Url"/>
{ribbon}{context}        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
{images}      </bt:Images>
      <bt:Urls>
{url_xml}      </bt:Urls>
      <bt:ShortStrings>
{short_xml}      </bt:ShortStrings>
      <bt:LongStrings>
{long_xml}      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
"""
root = Path(__file__).resolve().parent.parent

# ---- Shared runtime and keyboard shortcuts ----
# Word allows add-in keyboard shortcuts only in a shared runtime: ribbon, context
# menu and shortcuts all run as functions inside the task pane page
# (src/addin/actions.ts) instead of reloading it with ?action=…/?view=….
import re
import json

# Three-key shortcuts (Alt+Shift+letter): Ctrl+Alt+letter is AltGr on a Turkish
# keyboard (@, €, ₺ …) and many Ctrl+Shift letters are Word's own. On a Mac,
# Option+Shift types characters, so Cmd+Option there. Each is also written in
# the button's hover text. Alt+Shift+H, L, Y and Z are Word's own in Turkish Word.
SHORTCUTS = [
    ("go_action_insert", "Ölüm tarihi ekle", "Insert.Tip", "V"),
    ("go_action_spelling", "İsim yazımı", "Spelling.Tip", "W"),
    ("go_view_denetim", "Tüm belgeyi denetle", "Audit.Tip", "B"),
    ("go_view_dokum", "Şahıs dökümü", "Inventory.Tip", "G"),
    ("go_view_toplu", "Toplu düzeltme önizlemesi", "Bulk.Tip", "U"),
    ("go_view_ara", "Şahıs ara", "Search.Tip", "J"),
]

def go(match):
    query = urls[match.group(2)].split("?", 1)[1]
    key, value = query.split("=")
    return f'{match.group(1)}<Action xsi:type="ExecuteFunction">\n{match.group(1)}  <FunctionName>go_{key}_{value}</FunctionName>\n{match.group(1)}</Action>'

shortcut_item = item("Shortcuts", S("Shortcuts", "Klavye kısayolları"), L("Shortcuts.Tip", "Vefeyât kısayollarını gösterir; çakışan bir kısayolu boştaki bir tuşa alır."), "settings", pane(U("U.Shortcuts", "taskpane.html?view=kisayol")))
noproof_end = xml.index("</Item>\n", xml.index('id="Vefeyat.NoProof"')) + len("</Item>\n")
final = re.sub(r'( *)<Action xsi:type="ShowTaskpane">\s*<TaskpaneId>VefeyatPane</TaskpaneId>\s*<SourceLocation resid="([\w.]+)"/>\s*</Action>', go, xml[:noproof_end] + shortcut_item + xml[noproof_end:])
assert "ShowTaskpane" not in final
final = (final
    .replace('      <Set Name="WordApi" MinVersion="1.3"/>\n', '      <Set Name="WordApi" MinVersion="1.3"/>\n      <Set Name="SharedRuntime" MinVersion="1.1"/>\n')
    .replace('      <Host xsi:type="Document">\n', '      <Host xsi:type="Document">\n        <Runtimes>\n          <Runtime resid="Taskpane.Url" lifetime="long"/>\n        </Runtimes>\n')
    .replace('<FunctionFile resid="Commands.Url"/>', '<FunctionFile resid="Taskpane.Url"/>')
    .replace('        <bt:Url id="Commands.Url"', f'        <bt:Url id="Taskpane.Url" DefaultValue="{B}taskpane.html"/>\n        <bt:Url id="Commands.Url"')
    .replace("  </VersionOverrides>\n", f'  </VersionOverrides>\n  <ExtendedOverrides Url="{B}shortcuts.json"/>\n'))
final = final.replace("      </bt:ShortStrings>", f'        <bt:String id="Shortcuts" DefaultValue="{esc(short["Shortcuts"])}"/>\n      </bt:ShortStrings>')
final = final.replace("      </bt:LongStrings>", f'        <bt:String id="Shortcuts.Tip" DefaultValue="{esc(long_["Shortcuts.Tip"])}"/>\n      </bt:LongStrings>')
for action, name, tip, key in SHORTCUTS:
    assert f"<FunctionName>{action}</FunctionName>" in final, action
    old = f'<bt:String id="{tip}" DefaultValue="{esc(long_[tip])}"/>'
    assert old in final, tip
    final = final.replace(old, f'<bt:String id="{tip}" DefaultValue="Kısayol: Alt+Shift+{key}. {esc(long_[tip])}"/>')
    label = tip.removesuffix(".Tip")
    final = final.replace(f'<Title resid="{label}"/>', f'<Title resid="{label}.K"/>')
    final = final.replace("      </bt:ShortStrings>", f'        <bt:String id="{label}.K" DefaultValue="{esc(short[label])} (Alt+Shift+{key})"/>\n      </bt:ShortStrings>')
shortcuts = {
    "actions": [{"id": a, "type": "ExecuteFunction", "name": n} for a, n, _, _ in SHORTCUTS],
    "shortcuts": [{"action": a, "key": {"default": f"Alt+Shift+{k}", "mac": f"Command+Option+{k}"}} for a, _, _, k in SHORTCUTS],
}
root.joinpath("manifest/shortcuts.json").write_text(json.dumps(shortcuts, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
for needle in ("SharedRuntime", "<Runtimes>", 'FunctionFile resid="Taskpane.Url"', 'id="Taskpane.Url"', "ExtendedOverrides"):
    assert needle in final, needle
root.joinpath("manifest/manifest.template.xml").write_text(final, encoding="utf-8")
print(f"{len(used_icons)} icons, {len(urls)} urls, {len(short)} labels, {final.count('<FunctionName>go_')} ribbon functions, {len(SHORTCUTS)} shortcuts")
