from __future__ import annotations

import html
import zipfile
from pathlib import Path


OUT = Path("/Users/okadaigo/Documents/New project/PEC1-4_材料課題_簡易まとめ.pptx")

SLIDES = [
    (
        "PEC1-4のリサーチ結果と材料課題",
        ["卒業研究向け 簡易まとめ", "2026-04-16"],
    ),
    (
        "1. PEC1-4で技術として何をしたいか",
        [
            "PEC-1: ネットワーク側の光化",
            "データセンター間・サーバ間を低遅延・大容量・低消費電力でつなぎたい",
            "",
            "PEC-2: ボード間の光化",
            "Switch ASIC近傍まで optical engine を寄せて、装置内の通信電力を下げたい",
            "",
            "PEC-3: パッケージ間の光化",
            "半導体パッケージ間を高密度・低消費電力でつなぎたい",
            "",
            "PEC-4: ダイ間・チップ内の光化",
            "die-to-die / intra-chip まで光を入れて、計算機内部構造を変えたい",
        ],
    ),
    (
        "2. PEC-1 / PEC-2",
        [
            "PEC-1でやりたいこと",
            "長距離伝送を低損失・大容量で成立させる",
            "想定材料: 石英ファイバ、シリカ PLC / AWG、Si フォトニクス、InP 系",
            "材料課題: 伝送損失、波長分離、耐湿・耐熱、長期安定性",
            "",
            "PEC-2でやりたいこと",
            "ボード間配線を光化し、装置内通信の電力を下げる",
            "想定材料: Si / SiO2 / InP / InGaAlAs、SiOx 導波路、コネクタ・接着材",
            "材料課題: 熱、接合、アライメント、コネクタ小型化、近接実装",
        ],
    ),
    (
        "3. PEC-3",
        [
            "PEC-3でやりたいこと",
            "package-to-package 光接続を成立させる",
            "",
            "材料スタックの見方",
            "Si フォトニクス光回路 / InP 系膜型活性層 / クラッド / 放熱層 / 接着・封止材",
            "",
            "材料課題",
            "熱で光結合が変わる",
            "反りでアライメントがずれる",
            "1310 / 1550nm 損失が効く",
            "リフロー・熱サイクル・量産性が支配的",
            "",
            "研究の主戦場",
            "熱・反り・光損失・信頼性を同時に見る必要がある",
        ],
    ),
    (
        "4. PEC-4",
        [
            "PEC-4でやりたいこと",
            "die-to-die / chip内 optical layer を成立させる",
            "",
            "仮説的な材料スタック",
            "Si / SiN 導波路、InP 薄膜 or Ge / Si、低k / クラッド、Cu / SiCN bonding、熱拡散材",
            "",
            "材料課題",
            "bonding の位置ずれ・接合信頼性",
            "optical layer と electrical layer の熱干渉",
            "長期信頼性、応力、封止",
            "",
            "見立て",
            "全面光化より、I/O タイル起点の部分光化が自然そう",
        ],
    ),
    (
        "5. MGC材料への引きつけ",
        [
            "OPE",
            "高温耐性・CTE は比較的強い",
            "ただし 1310 / 1550nm の導波路損失が不明",
            "",
            "BT樹脂系",
            "基板材料として有望",
            "ただし反り・剥離・光学損失データが不足",
            "",
            "ユピゼータEP",
            "光学部材・光学樹脂候補",
            "ただし耐熱・近赤外・CTE の定量データが不足",
        ],
    ),
    (
        "6. 研究としての着地点",
        [
            "技術としてやりたいこと",
            "光電融合を、よりコンピュータの近くまで持ち込む",
            "",
            "材料研究としての問い",
            "どの材料が PEC-3 のどの役割を担えるか",
            "何のデータが足りないと最終判断できないか",
            "",
            "今後の評価軸",
            "260℃リフロー耐性 / 1310・1550nm 損失 / CTE・寸法安定性",
            "",
            "次にやること",
            "公開エビデンス比較と、評価項目-測定方法の対応づけ",
        ],
    ),
]


def xml_escape(text: str) -> str:
    return html.escape(text, quote=False)


def paragraph_xml(text: str, level: int = 0) -> str:
    if text == "":
        return "<a:p/>"
    return (
        f'<a:p><a:pPr lvl="{level}"/>'
        f'<a:r><a:rPr lang="ja-JP" sz="2200"/><a:t>{xml_escape(text)}</a:t></a:r>'
        f"</a:p>"
    )


def title_paragraph_xml(text: str) -> str:
    return (
        "<a:p>"
        f'<a:r><a:rPr lang="ja-JP" sz="2800" b="1"/><a:t>{xml_escape(text)}</a:t></a:r>'
        "</a:p>"
    )


def slide_xml(title: str, lines: list[str]) -> str:
    body_paragraphs = []
    for line in lines:
        if line == "":
            body_paragraphs.append(paragraph_xml(""))
        elif line.endswith("こと") or line in {
            "PEC-1: ネットワーク側の光化",
            "PEC-2: ボード間の光化",
            "PEC-3: パッケージ間の光化",
            "PEC-4: ダイ間・チップ内の光化",
            "材料スタックの見方",
            "材料課題",
            "研究の主戦場",
            "仮説的な材料スタック",
            "見立て",
            "OPE",
            "BT樹脂系",
            "ユピゼータEP",
            "技術としてやりたいこと",
            "材料研究としての問い",
            "今後の評価軸",
            "次にやること",
        }:
            body_paragraphs.append(paragraph_xml(line, 0))
        else:
            body_paragraphs.append(paragraph_xml(line, 1))
    body = "".join(body_paragraphs)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="0" y="0"/>
          <a:ext cx="0" cy="0"/>
          <a:chOff x="0" y="0"/>
          <a:chExt cx="0" cy="0"/>
        </a:xfrm>
      </p:grpSpPr>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="2" name="Title 1"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="457200" y="274320"/>
            <a:ext cx="8229600" cy="914400"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          {title_paragraph_xml(title)}
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr>
          <p:cNvPr id="3" name="Content Placeholder 2"/>
          <p:cNvSpPr/>
          <p:nvPr/>
        </p:nvSpPr>
        <p:spPr>
          <a:xfrm>
            <a:off x="640080" y="1341120"/>
            <a:ext cx="7924800" cy="4434840"/>
          </a:xfrm>
        </p:spPr>
        <p:txBody>
          <a:bodyPr wrap="square"/>
          <a:lstStyle/>
          {body}
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
  <p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>
"""


def slide_rel_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"""


def presentation_xml(num_slides: int) -> str:
    sld_ids = []
    for idx in range(1, num_slides + 1):
        sld_ids.append(
            f'<p:sldId id="{255 + idx}" r:id="rId{idx}"/>'
        )
    joined = "".join(sld_ids)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
 xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:sldMasterIdLst/>
  <p:sldIdLst>{joined}</p:sldIdLst>
  <p:sldSz cx="9144000" cy="5143500" type="screen16x9"/>
  <p:notesSz cx="6858000" cy="9144000"/>
  <p:defaultTextStyle/>
</p:presentation>
"""


def presentation_rels_xml(num_slides: int) -> str:
    rels = []
    for idx in range(1, num_slides + 1):
        rels.append(
            f'<Relationship Id="rId{idx}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{idx}.xml"/>'
        )
    joined = "".join(rels)
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  {joined}
</Relationships>
"""


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>
"""


ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>
"""


APP_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
 xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>OpenAI Codex</Application>
  <PresentationFormat>On-screen Show (16:9)</PresentationFormat>
  <Slides>6</Slides>
  <Notes>0</Notes>
  <HiddenSlides>0</HiddenSlides>
  <MMClips>0</MMClips>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant><vt:lpstr>Slides</vt:lpstr></vt:variant>
      <vt:variant><vt:i4>6</vt:i4></vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="6" baseType="lpstr">
      <vt:lpstr>Slide 1</vt:lpstr>
      <vt:lpstr>Slide 2</vt:lpstr>
      <vt:lpstr>Slide 3</vt:lpstr>
      <vt:lpstr>Slide 4</vt:lpstr>
      <vt:lpstr>Slide 5</vt:lpstr>
      <vt:lpstr>Slide 6</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <Company></Company>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>16.0000</AppVersion>
</Properties>
"""


CORE_XML = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:dcterms="http://purl.org/dc/terms/"
 xmlns:dcmitype="http://purl.org/dc/dcmitype/"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:title>PEC1-4のリサーチ結果と材料課題</dc:title>
  <dc:creator>OpenAI Codex</dc:creator>
  <cp:lastModifiedBy>OpenAI Codex</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2026-04-16T10:00:00Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2026-04-16T10:00:00Z</dcterms:modified>
</cp:coreProperties>
"""


def write_pptx() -> None:
    if OUT.exists():
        OUT.unlink()

    with zipfile.ZipFile(OUT, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", CONTENT_TYPES)
        zf.writestr("_rels/.rels", ROOT_RELS)
        zf.writestr("docProps/app.xml", APP_XML)
        zf.writestr("docProps/core.xml", CORE_XML)
        zf.writestr("ppt/presentation.xml", presentation_xml(len(SLIDES)))
        zf.writestr(
            "ppt/_rels/presentation.xml.rels",
            presentation_rels_xml(len(SLIDES)),
        )

        for idx, (title, lines) in enumerate(SLIDES, start=1):
            zf.writestr(f"ppt/slides/slide{idx}.xml", slide_xml(title, lines))
            zf.writestr(f"ppt/slides/_rels/slide{idx}.xml.rels", slide_rel_xml())


if __name__ == "__main__":
    write_pptx()
    print(OUT)
