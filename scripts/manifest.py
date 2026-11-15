"""Generates manifest/manifest.template.xml: the Vefeyât ribbon in the VBA 0.3.41 layout.

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


def soon(key, label, tip, ic, as_item=False):
    make = item if as_item else button
    return make(f"Soon.{key}", S(f"Soon.{key}", label), L(f"Soon.{key}.Tip", tip), ic, pane(U(f"U.Soon.{key}", f"taskpane.html?soon={key}")))


def setting_menu(key, label, tip, ic, choices):
    tip_id = L(f"S.{key}.Tip", tip + " Seçili değer Vefeyât penceresinin üstünde yazar.")
    items = [item(f"S.{key}.{v}", S(f"S.{key}.{v}", text), tip_id, ic, function(f"set_{key}_{v}")) for v, text in choices]
    return menu(f"S.{key}", S(f"S.{key}", label), tip_id, ic, items)


insert = button("Insert", S("Insert", "Ölüm tarihi ekle"),
                L("Insert.Tip", "Seçili adın ölüm tarihini ekler. Tek ve kesin eşleşmede, ilk geçişte ve adın arkasında tarih yoksa doğrudan ekler; aksi hâlde Vefeyât penceresinde seçim veya onay ister."),
                "insert", pane(U("U.Insert", "taskpane.html?action=insert")))
spelling = button("Spelling", S("Spelling", "İsim yazımı"),
                  L("Spelling.Tip", "Seçili adı kısa, tam ve orijinal adlarla karşılaştırır; Türkçe eki koruyarak yalnız seçili adı değiştirir."),
                  "spelling", pane(U("U.Spelling", "taskpane.html?action=spelling")))

groups = [
    group("Add", "Ekle ve denetle", "insert", [
        insert,
        button("Audit", S("Audit", "Tüm belgeyi denetle"), L("Audit.Tip", "Belgedeki adları sırayla gösterir; onaylayıp ilerlersiniz."), "audit", pane(U("U.Audit", "taskpane.html?view=denetim"))),
        spelling,
    ]),
    group("Review", "Denetim", "scope", [
        setting_menu("scope", "Tarih ekleme", "İlk kullanım (varsayılan), bütün kullanımlar veya yalnız denetim.", "scope",
                     [("first", "Yalnız ilk kullanım"), ("all", "Bütün kullanımlar"), ("none", "Yalnız denetle")]),
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
            soon("refresh", "Tarihleri güncelle", "Toplu önerileri önizlemede gösterir; mevcut tarihleri yeniden yazmaz.", "more", True),
            soon("compare", "Katalogla karşılaştır", "Belgedeki tarihleri güncel katalogla karşılaştırır.", "more", True),
            soon("registry", "Dinamik şahıs kayıtları", "Geçişleri kaynak kimliğine bağlar.", "more", True),
            soon("coverage", "Denetim kapsamını göster", "Taranan ve taranmayan bölümleri raporlar.", "more", True),
            soon("identity", "Kimlik eşleşmelerini düzelt", "Belirsiz adlar için verdiğiniz onayları yönetir.", "more", True),
            soon("exclusions", "Belge istisnaları", "Şahıs değil olarak işaretlenen ifadeleri yönetir.", "more", True),
        ]),
        button("Health", S("Health", "Belge sağlığı"), L("Health.Tip", "Vefeyât tarihlerini katalog ve önek ayarıyla karşılaştırır; yalnız seçtiklerinizi onarır."), "health", pane(U("U.Health", "taskpane.html?view=saglik"))),
        button("History", S("History", "İşlem geçmişi"), L("History.Tip", "Bu oturumdaki Vefeyât işlemlerini listeler; eklenen tarihlere gidebilirsiniz."), "history", pane(U("U.History", "taskpane.html?view=gecmis"))),
    ]),
    group("Submit", "Teslim ve dizin", "tools", [
        button("Consistency", S("Consistency", "Tutarlılık raporu"), L("Consistency.Tip", "ö./öl./v. ve tarih biçimlerinin karışık kullanıldığı yerleri gösterir; belgeyi değiştirmez."), "consistency", pane(U("U.Consistency", "taskpane.html?tool=tutarlilik"))),
        button("Report", S("Report", "Şahıs raporu"), L("Report.Tip", "Şahısları ilk geçiş sırasıyla ve kronolojik olarak yeni bir belgede listeler."), "report", pane(U("U.Report", "taskpane.html?tool=rapor"))),
        menu("Index", S("Index", "Şahıslar dizini"), L("Index.Tip", "Kimliği kesin geçişlerden Word şahıslar dizini."), "index", [
            item("IndexBuild", S("IndexBuild", "Dizini oluştur / güncelle"), "Index.Tip", "index", pane(U("U.IndexBuild", "taskpane.html?tool=dizin"))),
            item("IndexRemove", S("IndexRemove", "Dizin işaretlerini kaldır"), L("IndexRemove.Tip", "Yalnız Vefeyât dizin işaretlerini ve dizinini siler."), "index", pane(U("U.IndexRemove", "taskpane.html?tool=dizinkaldir"))),
        ]),
        button("Finalize", S("Finalize", "Teslim öncesi temizle"), L("Finalize.Tip", "Teslim kopyası için Vefeyât izlerini kaldırır; görünen metin değişmez."), "finalize", pane(U("U.Finalize", "taskpane.html?tool=temizlik"))),
    ]),
    group("Settings", "Tercihler", "nameform", [
        setting_menu("nameForm", "Ad biçimi", "İsim yazımında önce önerilecek biçim.", "nameform",
                     [("ask", "Her seferinde sor"), ("near", "Seçili ada yakın"), ("short", "Kısa ad"), ("full", "Tam / uzun ad"), ("original", "Orijinal ad")]),
        setting_menu("prefix", "Kısaltma", "Ölüm tarihinin başındaki kısaltma.", "prefix",
                     [("original", "Kaynakta olduğu gibi"), ("o", "ö."), ("ol", "öl."), ("v", "v.")]),
        menu("Prefs", S("Prefs", "Ayarlar"), L("Prefs.Tip", "Yazım tercihleri ve görünüm ayarları."), "settings", [
            soon("preferSpelling", "Seçili yazımı tercih et", "Seçili yazımı bu şahıs için tercih olarak kaydeder.", "settings", True),
            soon("clearSpelling", "Yazım tercihlerini temizle", "Belgeye özgü yazım tercihlerini siler.", "settings", True),
            soon("exportPrefs", "Tercihleri dışa aktar", "Tercihleri bir dosyaya kaydeder.", "settings", True),
            soon("importPrefs", "Tercihleri al", "Tercih dosyasını bu belgeye uygular.", "settings", True),
            soon("shading", "Alan gölgelerini gizle", "Tarih alanlarının gri gölgesini gizler.", "settings", True),
        ]),
    ]),
    group("Data", "Veri ve yardım", "info", [
        button("Info", S("Info", "Veri bilgisi"), L("Info.Tip", "Yüklü kataloğu ve sürümü gösterir."), "info", pane(U("U.Info", "taskpane.html?view=katalog"))),
        button("Csv", S("Csv", "CSV seç"), L("Csv.Tip", "Bilgisayarınızdaki başka bir DİA CSV dosyasını kullanır."), "csv", pane(U("U.Csv", "taskpane.html?view=katalog"))),
        button("Help", S("Help", "Kullanım kılavuzu"), L("Help.Tip", "Vefeyât komutlarının kısa kılavuzu."), "help", pane(U("U.Help", "taskpane.html?view=yardim"))),
    ]),
]

ribbon = ("          <ExtensionPoint xsi:type=\"PrimaryCommandSurface\">\n            <CustomTab id=\"Vefeyat.Tab\">\n"
          + "".join(groups) + f'              <Label resid="{S("Tab", "Vefeyât")}"/>\n            </CustomTab>\n          </ExtensionPoint>\n')
context = ("          <ExtensionPoint xsi:type=\"ContextMenu\">\n            <OfficeMenu id=\"ContextMenuText\">\n"
           + menu("Context", S("Context", "Vefeyât"), L("Context.Tip", "Seçili şahıs adı için Vefeyât komutları."), "insert", [
               item("C.Insert", "Insert", "Insert.Tip", "insert", pane("U.Insert")),
               item("C.Spelling", "Spelling", "Spelling.Tip", "spelling", pane("U.Spelling")),
               item("C.Dia", "Dia", "Dia.Tip", "dia", pane("U.Dia")),
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
  <ProviderName>Vefeyât</ProviderName>
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
Path(__file__).resolve().parent.parent.joinpath("manifest/manifest.template.xml").write_text(xml, encoding="utf-8")
print(f"{len(used_icons)} icons, {len(urls)} urls, {len(short)} labels")
