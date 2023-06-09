import JSZip from 'jszip'
import CellStyle from './CellStyle.js'
import getColumnNameByIndex from './getColumnNameByIndex.js'

export default async function exportToExcel(streamer) { // returns Blob

    const fontName = 'Calibri'
    const fontSize = 11

    let strings = []
    let count = 0
    let rowNo = 1
    let result = ''
    const dimensions = {
        cols: 0,
        rows: 0
    }
    const columnsExtrems = []

    const styles = [ new CellStyle() ]
    const getStyleId = style => {
        let i = styles.findIndex(item => item.isEquals(style))
        if (i < 0) {
            i = styles.push(style) - 1
        }
        return i
    }

    streamer.streamAll((values, styles) => {
        const row = []
        let colNo = 0
        values.forEach((value, i) => {
            const styleId = getStyleId(styles[i])
            let stringId = strings.indexOf(value)
            if (stringId < 0) {
                stringId = strings.push(value) - 1
            }
            row.push(`<c r="${ getColumnNameByIndex(colNo) + rowNo }" s="${ styleId }" t="s"><v>${ stringId }</v></c>`)
            if (dimensions.cols < colNo) {
                dimensions.cols = colNo
            }
            if (columnsExtrems[colNo] === undefined || columnsExtrems[colNo].value.length < value.length) {
                columnsExtrems[colNo] = { value }
            }
            colNo ++
            count ++
        })
        result += `<row r="${ rowNo }" spans="1:${row.length}">\n${ row.join('\n') }\n</row>\n`
        rowNo ++
        dimensions.rows ++
    })

    columnsExtrems.forEach(col => {
        col.width = col.value.length + 2
    })
    
    strings = '<?xml version="1.0" encoding="utf-8" standalone="yes"?>'
    + `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${ count }" uniqueCount="${ strings.length }">\n`
    + strings.map(value => `<si><t>${ value }</t></si>`).join('\n')
    + '\n</sst>'
    // console.log(strings)

    const root = new JSZip()

    const _rels = root.folder('_rels')
    const docProps = root.folder('docProps')
    const xl = root.folder('xl')
    const xl__rels = xl.folder('_rels')
    const xl_theme = xl.folder('theme')
    const xl_worksheets = xl.folder('worksheets')

    root.file('[Content_Types].xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" />
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" />
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
  <Default Extension="xml" ContentType="application/xml" />
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" />
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" />
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" />
  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" />
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml" />
</Types>`)

    _rels.file('.rels', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml" />
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml" />
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml" />
</Relationships>`)

    docProps.file('app.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>SheetJS</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs>
    <vt:vector size="2" baseType="variant">
      <vt:variant>
        <vt:lpstr>Листы</vt:lpstr>
      </vt:variant>
      <vt:variant>
        <vt:i4>1</vt:i4>
      </vt:variant>
    </vt:vector>
  </HeadingPairs>
  <TitlesOfParts>
    <vt:vector size="1" baseType="lpstr">
      <vt:lpstr>Лист1</vt:lpstr>
    </vt:vector>
  </TitlesOfParts>
  <LinksUpToDate>false</LinksUpToDate>
  <SharedDoc>false</SharedDoc>
  <HyperlinksChanged>false</HyperlinksChanged>
  <AppVersion>12.0000</AppVersion>
</Properties>`)

    docProps.file('core.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>azbitnev</dc:creator>
  <cp:lastModifiedBy>azbitnev</cp:lastModifiedBy>
  <dcterms:created xsi:type="dcterms:W3CDTF">2022-06-21T05:48:54Z</dcterms:created>
  <dcterms:modified xsi:type="dcterms:W3CDTF">2022-06-21T05:48:54Z</dcterms:modified>
</cp:coreProperties>`)

    xl__rels.file('workbook.xml.rels', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml" />
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml" />
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml" />
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml" />
</Relationships>`)

    xl_theme.file('theme1.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme">
  <a:themeElements>
    <a:clrScheme name="Office">
      <a:dk1>
        <a:sysClr val="windowText" lastClr="000000" />
      </a:dk1>
      <a:lt1>
        <a:sysClr val="window" lastClr="FFFFFF" />
      </a:lt1>
      <a:dk2>
        <a:srgbClr val="1F497D" />
      </a:dk2>
      <a:lt2>
        <a:srgbClr val="EEECE1" />
      </a:lt2>
      <a:accent1>
        <a:srgbClr val="4F81BD" />
      </a:accent1>
      <a:accent2>
        <a:srgbClr val="C0504D" />
      </a:accent2>
      <a:accent3>
        <a:srgbClr val="9BBB59" />
      </a:accent3>
      <a:accent4>
        <a:srgbClr val="8064A2" />
      </a:accent4>
      <a:accent5>
        <a:srgbClr val="4BACC6" />
      </a:accent5>
      <a:accent6>
        <a:srgbClr val="F79646" />
      </a:accent6>
      <a:hlink>
        <a:srgbClr val="0000FF" />
      </a:hlink>
      <a:folHlink>
        <a:srgbClr val="800080" />
      </a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Office">
      <a:majorFont>
        <a:latin typeface="Cambria" />
        <a:ea typeface="" />
        <a:cs typeface="" />
        <a:font script="Jpan" typeface="ＭＳ Ｐゴシック" />
        <a:font script="Hang" typeface="맑은 고딕" />
        <a:font script="Hans" typeface="宋体" />
        <a:font script="Hant" typeface="新細明體" />
        <a:font script="Arab" typeface="Times New Roman" />
        <a:font script="Hebr" typeface="Times New Roman" />
        <a:font script="Thai" typeface="Tahoma" />
        <a:font script="Ethi" typeface="Nyala" />
        <a:font script="Beng" typeface="Vrinda" />
        <a:font script="Gujr" typeface="Shruti" />
        <a:font script="Khmr" typeface="MoolBoran" />
        <a:font script="Knda" typeface="Tunga" />
        <a:font script="Guru" typeface="Raavi" />
        <a:font script="Cans" typeface="Euphemia" />
        <a:font script="Cher" typeface="Plantagenet Cherokee" />
        <a:font script="Yiii" typeface="Microsoft Yi Baiti" />
        <a:font script="Tibt" typeface="Microsoft Himalaya" />
        <a:font script="Thaa" typeface="MV Boli" />
        <a:font script="Deva" typeface="Mangal" />
        <a:font script="Telu" typeface="Gautami" />
        <a:font script="Taml" typeface="Latha" />
        <a:font script="Syrc" typeface="Estrangelo Edessa" />
        <a:font script="Orya" typeface="Kalinga" />
        <a:font script="Mlym" typeface="Kartika" />
        <a:font script="Laoo" typeface="DokChampa" />
        <a:font script="Sinh" typeface="Iskoola Pota" />
        <a:font script="Mong" typeface="Mongolian Baiti" />
        <a:font script="Viet" typeface="Times New Roman" />
        <a:font script="Uigh" typeface="Microsoft Uighur" />
        <a:font script="Geor" typeface="Sylfaen" />
      </a:majorFont>
      <a:minorFont>
        <a:latin typeface="Calibri" />
        <a:ea typeface="" />
        <a:cs typeface="" />
        <a:font script="Jpan" typeface="ＭＳ Ｐゴシック" />
        <a:font script="Hang" typeface="맑은 고딕" />
        <a:font script="Hans" typeface="宋体" />
        <a:font script="Hant" typeface="新細明體" />
        <a:font script="Arab" typeface="Arial" />
        <a:font script="Hebr" typeface="Arial" />
        <a:font script="Thai" typeface="Tahoma" />
        <a:font script="Ethi" typeface="Nyala" />
        <a:font script="Beng" typeface="Vrinda" />
        <a:font script="Gujr" typeface="Shruti" />
        <a:font script="Khmr" typeface="DaunPenh" />
        <a:font script="Knda" typeface="Tunga" />
        <a:font script="Guru" typeface="Raavi" />
        <a:font script="Cans" typeface="Euphemia" />
        <a:font script="Cher" typeface="Plantagenet Cherokee" />
        <a:font script="Yiii" typeface="Microsoft Yi Baiti" />
        <a:font script="Tibt" typeface="Microsoft Himalaya" />
        <a:font script="Thaa" typeface="MV Boli" />
        <a:font script="Deva" typeface="Mangal" />
        <a:font script="Telu" typeface="Gautami" />
        <a:font script="Taml" typeface="Latha" />
        <a:font script="Syrc" typeface="Estrangelo Edessa" />
        <a:font script="Orya" typeface="Kalinga" />
        <a:font script="Mlym" typeface="Kartika" />
        <a:font script="Laoo" typeface="DokChampa" />
        <a:font script="Sinh" typeface="Iskoola Pota" />
        <a:font script="Mong" typeface="Mongolian Baiti" />
        <a:font script="Viet" typeface="Arial" />
        <a:font script="Uigh" typeface="Microsoft Uighur" />
        <a:font script="Geor" typeface="Sylfaen" />
      </a:minorFont>
    </a:fontScheme>
    <a:fmtScheme name="Office">
      <a:fillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr" />
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="50000" />
                <a:satMod val="300000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="35000">
              <a:schemeClr val="phClr">
                <a:tint val="37000" />
                <a:satMod val="300000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="15000" />
                <a:satMod val="350000" />
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="1" />
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="100000" />
                <a:shade val="100000" />
                <a:satMod val="130000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:tint val="50000" />
                <a:shade val="100000" />
                <a:satMod val="350000" />
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:lin ang="16200000" scaled="0" />
        </a:gradFill>
      </a:fillStyleLst>
      <a:lnStyleLst>
        <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr">
              <a:shade val="95000" />
              <a:satMod val="105000" />
            </a:schemeClr>
          </a:solidFill>
          <a:prstDash val="solid" />
        </a:ln>
        <a:ln w="25400" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr" />
          </a:solidFill>
          <a:prstDash val="solid" />
        </a:ln>
        <a:ln w="38100" cap="flat" cmpd="sng" algn="ctr">
          <a:solidFill>
            <a:schemeClr val="phClr" />
          </a:solidFill>
          <a:prstDash val="solid" />
        </a:ln>
      </a:lnStyleLst>
      <a:effectStyleLst>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="38000" />
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000" />
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
        </a:effectStyle>
        <a:effectStyle>
          <a:effectLst>
            <a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0">
              <a:srgbClr val="000000">
                <a:alpha val="35000" />
              </a:srgbClr>
            </a:outerShdw>
          </a:effectLst>
          <a:scene3d>
            <a:camera prst="orthographicFront">
              <a:rot lat="0" lon="0" rev="0" />
            </a:camera>
            <a:lightRig rig="threePt" dir="t">
              <a:rot lat="0" lon="0" rev="1200000" />
            </a:lightRig>
          </a:scene3d>
          <a:sp3d>
            <a:bevelT w="63500" h="25400" />
          </a:sp3d>
        </a:effectStyle>
      </a:effectStyleLst>
      <a:bgFillStyleLst>
        <a:solidFill>
          <a:schemeClr val="phClr" />
        </a:solidFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="40000" />
                <a:satMod val="350000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="40000">
              <a:schemeClr val="phClr">
                <a:tint val="45000" />
                <a:shade val="99000" />
                <a:satMod val="350000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="20000" />
                <a:satMod val="255000" />
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="-80000" r="50000" b="180000" />
          </a:path>
        </a:gradFill>
        <a:gradFill rotWithShape="1">
          <a:gsLst>
            <a:gs pos="0">
              <a:schemeClr val="phClr">
                <a:tint val="80000" />
                <a:satMod val="300000" />
              </a:schemeClr>
            </a:gs>
            <a:gs pos="100000">
              <a:schemeClr val="phClr">
                <a:shade val="30000" />
                <a:satMod val="200000" />
              </a:schemeClr>
            </a:gs>
          </a:gsLst>
          <a:path path="circle">
            <a:fillToRect l="50000" t="50000" r="50000" b="50000" />
          </a:path>
        </a:gradFill>
      </a:bgFillStyleLst>
    </a:fmtScheme>
  </a:themeElements>
  <a:objectDefaults>
    <a:spDef>
      <a:spPr />
      <a:bodyPr />
      <a:lstStyle />
      <a:style>
        <a:lnRef idx="1">
          <a:schemeClr val="accent1" />
        </a:lnRef>
        <a:fillRef idx="3">
          <a:schemeClr val="accent1" />
        </a:fillRef>
        <a:effectRef idx="2">
          <a:schemeClr val="accent1" />
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="lt1" />
        </a:fontRef>
      </a:style>
    </a:spDef>
    <a:lnDef>
      <a:spPr />
      <a:bodyPr />
      <a:lstStyle />
      <a:style>
        <a:lnRef idx="2">
          <a:schemeClr val="accent1" />
        </a:lnRef>
        <a:fillRef idx="0">
          <a:schemeClr val="accent1" />
        </a:fillRef>
        <a:effectRef idx="1">
          <a:schemeClr val="accent1" />
        </a:effectRef>
        <a:fontRef idx="minor">
          <a:schemeClr val="tx1" />
        </a:fontRef>
      </a:style>
    </a:lnDef>
  </a:objectDefaults>
  <a:extraClrSchemeLst />
</a:theme>`)

    xl.file('workbook.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4506" />
  <workbookPr codeName="ThisWorkbook" />
  <bookViews>
    <workbookView xWindow="480" yWindow="855" windowWidth="28215" windowHeight="11670" />
  </bookViews>
  <sheets>
    <sheet name="Лист1" sheetId="1" r:id="rId1" />
  </sheets>
  <calcPr calcId="125725" />
</workbook>`)

    xl.file('styles.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="2">
    <font>
      <sz val="${ fontSize }" />
      <color theme="1" />
      <name val="${ fontName }" />
      <family val="2" />
      <scheme val="minor" />
    </font>
    <font>
      <b />
      <sz val="${ fontSize }" />
      <color theme="1" />
      <name val="${ fontName }" />
      <family val="2" />
      <charset val="204" />
      <scheme val="minor" />
    </font>
  </fonts>
  <fills count="2">
    <fill>
      <patternFill patternType="none" />
    </fill>
    <fill>
      <patternFill patternType="gray125" />
    </fill>
  </fills>
  <borders count="1">
    <border>
      <left />
      <right />
      <top />
      <bottom />
      <diagonal />
    </border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" />
  </cellStyleXfs>
  <cellXfs count="${ styles.length }">
    ${ styles.map(style => `<xf numFmtId="0" fontId="${ style.isBold ? 1 : 0 }" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"${ style.isBold ? ' applyFont="1"' : '' } applyAlignment="1">
      <alignment horizontal="${ style.horizontalAlignment }"/>
    </xf>`).join('\n') }
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Обычный" xfId="0" builtinId="0" />
  </cellStyles>
  <dxfs count="0" />
  <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleMedium4" />
</styleSheet>`)

    xl.file('sharedStrings.xml', strings)

    xl_worksheets.file('sheet1.xml', `<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <dimension ref="A1:${ getColumnNameByIndex(dimensions.cols) + dimensions.rows }" />
  <sheetViews>
    <sheetView tabSelected="1" workbookViewId="0">
      <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen" />
      <selection pane="bottomLeft" activeCell="A2" sqref="A2" />
    </sheetView>
  </sheetViews>
  <sheetFormatPr defaultRowHeight="15" />
  <cols>
    ${ columnsExtrems.map((col, i) => `<col min="${ i + 1 }" max="${ i + 1 }" width="${ col.width }" customWidth="1" />`).join('\n') }
  </cols>
  <sheetData>
    ${ result }
  </sheetData>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" />
</worksheet>`)

    const blob = await root.generateAsync({ type: 'blob' })

    return blob
}
