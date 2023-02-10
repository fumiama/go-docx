package docxlib

import (
	"encoding/xml"
	"io"
	"os"
	"testing"
)

const decoded_doc_1 = `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex" xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex" xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex" xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex" xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex" xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex" xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex" xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink" xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex" xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid" xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex wp14"><w:body><w:p w14:paraId="77CA082D" w14:textId="4AF3264D" w:rsidR="00D66E3F" w:rsidRDefault="003A3F42"><w:pPr><w:rPr><w:color w:val="808080"/></w:rPr></w:pPr><w:proofErr w:type="spellStart"/><w:r><w:t>test</w:t></w:r><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>test</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t xml:space="preserve"> font </w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:rPr><w:sz w:val="44"/></w:rPr><w:t>size</w:t></w:r><w:r><w:rPr><w:color w:val="808080"/></w:rPr><w:t>test</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r><w:rPr><w:color w:val="808080"/></w:rPr><w:t xml:space="preserve"> color</w:t></w:r></w:p><w:p w14:paraId="6D114165" w14:textId="04580C29" w:rsidR="003A3F42" w:rsidRDefault="003A3F42" w:rsidP="003A3F42"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>New style 1</w:t></w:r></w:p><w:p w14:paraId="40D72B3B" w14:textId="76101901" w:rsidR="003A3F42" w:rsidRDefault="003A3F42" w:rsidP="003A3F42"><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>New style 2</w:t></w:r></w:p><w:p w14:paraId="1CA8A9B3" w14:textId="77777777" w:rsidR="00D66E3F" w:rsidRDefault="003A3F42"><w:r><w:rPr><w:color w:val="FF0000"/><w:sz w:val="44"/></w:rPr><w:t>test font size and color</w:t></w:r></w:p><w:p w14:paraId="0D82FB8B" w14:textId="77777777" w:rsidR="00D66E3F" w:rsidRDefault="003A3F42"><w:hyperlink r:id="rId4"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>google</w:t></w:r></w:hyperlink></w:p><w:sectPr w:rsidR="00D66E3F"><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`
const decoded_doc_2 = `<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 wp14"><w:body><w:sdt><w:sdtPr><w:id w:val="-1247033294"/><w:docPartObj><w:docPartGallery w:val="Table of Contents"/><w:docPartUnique/></w:docPartObj></w:sdtPr><w:sdtEndPr/><w:sdtContent><w:p w14:paraId="308E3D65" w14:textId="77777777" w:rsidR="00FA66BB" w:rsidRPr="001D59EF" w:rsidRDefault="00FA66BB" w:rsidP="00A96827"><w:pPr><w:pStyle w:val="TOC1"/><w:jc w:val="center"/></w:pPr><w:r><w:t>Table of Contents</w:t></w:r></w:p></w:sdtContent></w:sdt><w:p w14:paraId="1764C163" w14:textId="77777777" w:rsidR="009E307C" w:rsidRDefault="00A96827"><w:pPr><w:pStyle w:val="TOC1"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:b w:val="0"/><w:color w:val="auto"/></w:rPr></w:pPr><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:instrText xml:space="preserve"> TOC \h \z \t "Heading 1,2,S6,1,S0,1,S1,1,S2,1,S3,1,S4,1,S5,1" </w:instrText></w:r><w:r><w:rPr><w:b w:val="0"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:hyperlink w:anchor="_Toc420414504" w:history="1"><w:r w:rsidR="009E307C" w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Holy Grail [xref:bRJduW6hNR]</w:t></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:tab/></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF _Toc420414504 \h </w:instrText></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:t>2</w:t></w:r><w:r w:rsidR="009E307C"><w:rPr><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p><w:p w14:paraId="0F5BA552" w14:textId="77777777" w:rsidR="009E307C" w:rsidRDefault="009E307C"><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="left" w:pos="3654"/></w:tabs><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr></w:pPr><w:hyperlink w:anchor="_Toc420414505" w:history="1"><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>1.</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr><w:tab/></w:r><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>What is your name? [xref:TH7u7QDqhD]</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF _Toc420414505 \h </w:instrText></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p><w:p w14:paraId="49E1F0AA" w14:textId="77777777" w:rsidR="009E307C" w:rsidRDefault="009E307C"><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="left" w:pos="3654"/></w:tabs><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr></w:pPr><w:hyperlink w:anchor="_Toc420414506" w:history="1"><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>2.</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr><w:tab/></w:r><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>What is your quest? [xref:bC62HkFATC]</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF _Toc420414506 \h </w:instrText></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p><w:p w14:paraId="7BDA743C" w14:textId="77777777" w:rsidR="009E307C" w:rsidRDefault="009E307C"><w:pPr><w:pStyle w:val="TOC2"/><w:tabs><w:tab w:val="left" w:pos="3654"/></w:tabs><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr></w:pPr><w:hyperlink w:anchor="_Toc420414507" w:history="1"><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>3.</w:t></w:r><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi"/><w:noProof/></w:rPr><w:tab/></w:r><w:r w:rsidRPr="002306B2"><w:rPr><w:rStyle w:val="Hyperlink"/><w:noProof/></w:rPr><w:t>What is your favourite colour? [xref:I3TphuHX6N]</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:tab/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:instrText xml:space="preserve"> PAGEREF _Toc420414507 \h </w:instrText></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:noProof/><w:webHidden/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:hyperlink></w:p><w:p w14:paraId="4A0A0E88" w14:textId="77777777" w:rsidR="006C0D12" w:rsidRPr="009C59B6" w:rsidRDefault="00A96827" w:rsidP="009B657F"><w:pPr><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p><w:p w14:paraId="7EDC60AD" w14:textId="77777777" w:rsidR="0004272B" w:rsidRDefault="0004272B"><w:pPr><w:jc w:val="left"/><w:rPr><w:b/></w:rPr></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:br w:type="page"/></w:r><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/></w:p><w:p w14:paraId="6775D4EA" w14:textId="4B1B4185" w:rsidR="00EF5BF6" w:rsidRDefault="00DE7E6E" w:rsidP="00EF5BF6"><w:pPr><w:pStyle w:val="S0"/></w:pPr><w:bookmarkStart w:id="1" w:name="_Toc388285991"/><w:bookmarkStart w:id="2" w:name="_Toc388366779"/><w:bookmarkStart w:id="3" w:name="_Toc388428327"/><w:bookmarkStart w:id="4" w:name="_Toc388451002"/><w:bookmarkStart w:id="5" w:name="_Toc420414504"/><w:r><w:lastRenderedPageBreak/><w:t>Holy Grail</w:t></w:r><w:bookmarkEnd w:id="1"/><w:bookmarkEnd w:id="2"/><w:bookmarkEnd w:id="3"/><w:bookmarkEnd w:id="4"/><w:r w:rsidR="009E307C"><w:t xml:space="preserve"> [</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="bookmark"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>xref</w:t></w:r><w:proofErr w:type="gramStart"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>:bRJduW6hNR</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r w:rsidR="009E307C"><w:t>]</w:t></w:r><w:bookmarkEnd w:id="5"/></w:p><w:p w14:paraId="2E760FD1" w14:textId="10909973" w:rsidR="00DE7E6E" w:rsidRDefault="00DE7E6E" w:rsidP="00DE7E6E"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="6" w:name="_Toc389482870"/><w:bookmarkStart w:id="7" w:name="_Toc420414505"/><w:r><w:t>What is your name?</w:t></w:r><w:bookmarkEnd w:id="6"/><w:r w:rsidR="009E307C"><w:t xml:space="preserve"> [</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="bookmark"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:proofErr w:type="gramStart"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>xref:</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>TH7u7QDqhD</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r w:rsidR="009E307C"><w:t>]</w:t></w:r><w:bookmarkEnd w:id="7"/></w:p><w:p w14:paraId="0249939B" w14:textId="77777777" w:rsidR="003946B5" w:rsidRPr="003946B5" w:rsidRDefault="00DE7E6E" w:rsidP="003946B5"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r w:rsidRPr="0029440C"><w:t xml:space="preserve">My name is Sir </w:t></w:r><w:proofErr w:type="spellStart"/><w:r w:rsidRPr="0029440C"><w:t>Launcelot</w:t></w:r><w:proofErr w:type="spellEnd"/><w:r w:rsidRPr="0029440C"><w:t xml:space="preserve"> of Camelot.</w:t></w:r></w:p><w:p w14:paraId="5BB04A25" w14:textId="04E09ADD" w:rsidR="006F5AAA" w:rsidRPr="006F5AAA" w:rsidRDefault="00DE7E6E" w:rsidP="00DE7E6E"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="8" w:name="_Toc389482871"/><w:bookmarkStart w:id="9" w:name="_Toc420414506"/><w:r><w:t>What is your quest?</w:t></w:r><w:bookmarkEnd w:id="8"/><w:r w:rsidR="009E307C"><w:t xml:space="preserve"> [</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="bookmark"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:proofErr w:type="gramStart"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>xref:</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>bC62HkFATC</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:r w:rsidR="009E307C"><w:t>]</w:t></w:r><w:bookmarkEnd w:id="9"/></w:p><w:p w14:paraId="15194710" w14:textId="77777777" w:rsidR="002B0891" w:rsidRDefault="00DE7E6E" w:rsidP="002B0891"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t xml:space="preserve">To seek the Holy </w:t></w:r><w:proofErr w:type="gramStart"/><w:r><w:t>Grail</w:t></w:r><w:r w:rsidRPr="00225D92"><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>[</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidRPr="00225D92"><w:rPr><w:color w:val="FF0000"/></w:rPr><w:t>or a grail shaped beacon]</w:t></w:r><w:r><w:t>.</w:t></w:r><w:r w:rsidR="00585075"><w:t xml:space="preserve"> </w:t></w:r></w:p><w:p w14:paraId="05C7DE39" w14:textId="2A77E45D" w:rsidR="00585075" w:rsidRDefault="00DE7E6E" w:rsidP="006F5AAA"><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:bookmarkStart w:id="10" w:name="_Toc389482872"/><w:bookmarkStart w:id="11" w:name="_Toc420414507"/><w:r><w:t>What is your favourite colour?</w:t></w:r><w:bookmarkEnd w:id="10"/><w:r w:rsidR="009E307C"><w:t xml:space="preserve"> [</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="bookmark"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:bookmarkStart w:id="12" w:name="bookmark"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:proofErr w:type="gramStart"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>xref:</w:t></w:r><w:proofErr w:type="gramEnd"/><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:t>I3TphuHX6N</w:t></w:r><w:r w:rsidR="009E307C" w:rsidRPr="009E307C"><w:rPr><w:b w:val="0"/><w:color w:val="808080"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r><w:bookmarkEnd w:id="12"/><w:r w:rsidR="009E307C"><w:t>]</w:t></w:r><w:bookmarkEnd w:id="11"/></w:p><w:p w14:paraId="5FA4E707" w14:textId="77777777" w:rsidR="00DE7E6E" w:rsidRDefault="00DE7E6E" w:rsidP="00DE7E6E"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>Blue.</w:t></w:r></w:p><w:p w14:paraId="543FEBD5" w14:textId="77777777" w:rsidR="006F5AAA" w:rsidRPr="006F5AAA" w:rsidRDefault="00DE7E6E" w:rsidP="00DE7E6E"><w:pPr><w:pStyle w:val="BodyText"/></w:pPr><w:r><w:t>How many paragraphs here then?</w:t></w:r></w:p><w:sectPr w:rsidR="006F5AAA" w:rsidRPr="006F5AAA" w:rsidSect="002B3068"><w:footerReference w:type="default" r:id="rId9"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="1134" w:bottom="1134" w:left="1134" w:header="709" w:footer="709" w:gutter="0"/><w:cols w:space="708"/><w:docGrid w:linePitch="360"/></w:sectPr></w:body></w:document>`

func TestUnmarshalPlainStructure(t *testing.T) {
	doc := Document{
		XMLW:    XMLNS_W,
		XMLR:    XMLNS_R,
		XMLWP:   XMLNS_WP,
		XMLName: xml.Name{Space: XMLNS_W, Local: "document"}}
	testCases := []struct {
		content       string
		numParagraphs int
	}{
		{decoded_doc_1, 5},
		{decoded_doc_2, 19},
	}
	for _, tc := range testCases {
		err := xml.Unmarshal(StringToBytes(tc.content), &doc)
		if err != nil {
			t.Fatal(err)
		}
		if len(doc.Body.Paragraphs) != tc.numParagraphs {
			t.Fatalf("We expected %d paragraphs, we got %d", tc.numParagraphs, len(doc.Body.Paragraphs))
		}
		for i, p := range doc.Body.Paragraphs {
			if len(p.Children) == 0 {
				t.Fatalf("We were not able to parse paragraph %d", i)
			}
			for _, child := range p.Children {
				if child.Link == nil && child.Properties == nil && child.Run == nil {
					t.Fatalf("There are Paragraph children with all fields nil")
				}
				if child.Run != nil && child.Run.Text == nil && child.Run.InstrText == "" {
					t.Fatalf("We have a run with no text")
				}
				if child.Link != nil && child.Link.ID == "" {
					t.Fatalf("We have a link without ID")
				}
			}
		}
	}
}

const drawing_doc = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas"
    xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
    xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex"
    xmlns:cx2="http://schemas.microsoft.com/office/drawing/2015/10/21/chartex"
    xmlns:cx3="http://schemas.microsoft.com/office/drawing/2016/5/9/chartex"
    xmlns:cx4="http://schemas.microsoft.com/office/drawing/2016/5/10/chartex"
    xmlns:cx5="http://schemas.microsoft.com/office/drawing/2016/5/11/chartex"
    xmlns:cx6="http://schemas.microsoft.com/office/drawing/2016/5/12/chartex"
    xmlns:cx7="http://schemas.microsoft.com/office/drawing/2016/5/13/chartex"
    xmlns:cx8="http://schemas.microsoft.com/office/drawing/2016/5/14/chartex"
    xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
    xmlns:aink="http://schemas.microsoft.com/office/drawing/2016/ink"
    xmlns:am3d="http://schemas.microsoft.com/office/drawing/2017/model3d"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    xmlns:oel="http://schemas.microsoft.com/office/2019/extlst"
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
    xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
    xmlns:w10="urn:schemas-microsoft-com:office:word"
    xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"
    xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"
    xmlns:w16cid="http://schemas.microsoft.com/office/word/2016/wordml/cid"
    xmlns:w16="http://schemas.microsoft.com/office/word/2018/wordml"
    xmlns:w16sdtdh="http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash"
    xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex"
    xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
    xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk"
    xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml"
    xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
    <w:body>
        <w:p w14:paraId="53D18053" w14:textId="034941A2" w:rsidR="00EE4CB0" w:rsidRDefault="00EE4CB0">
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
                <w:t>直接粘贴</w:t>
            </w:r>
            <w:r w:rsidR="00F71D77">
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
                <w:t xml:space="preserve"></w:t>
            </w:r>
            <w:r w:rsidR="00F71D77">
                <w:t>inline</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="525D8E7A" w14:textId="19079AFC" w:rsidR="00694680" w:rsidRDefault="00D04928">
            <w:r>
                <w:rPr>
                    <w:noProof/>
                </w:rPr>
                <w:drawing>
                    <wp:inline distT="T-mock-inline-p1-c0" distB="B-mock-inline-p1-c0" distL="L-mock-inline-p1-c0" distR="R-mock-inline-p1-c0" wp14:anchorId="3D4E5BAA" wp14:editId="3F7CEF85">
                        <wp:extent cx="5274310" cy="3369310"/>
                        <wp:effectExtent l="0" t="0" r="0" b="0"/>
                        <wp:docPr id="1" name="图片 1"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="1" name="图片 1"/>
                                        <pic:cNvPicPr/>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="rId4" cstate="print">
                                            <a:extLst>
                                                <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                    <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                                </a:ext>
                                            </a:extLst>
                                        </a:blip>
                                        <a:stretch>
                                            <a:fillRect/>
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:off x="0" y="0"/>
                                            <a:ext cx="5274310" cy="3369310"/>
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst/>
                                        </a:prstGeom>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
        <w:p w14:paraId="272A3387" w14:textId="6FF6E032" w:rsidR="00EE4CB0" w:rsidRDefault="00A549AD">
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
                <w:t xml:space="preserve">一行2个 </w:t>
            </w:r>
            <w:r>
                <w:t>inline</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="03284AA8" w14:textId="5134D375" w:rsidR="00EE4CB0" w:rsidRDefault="00EE4CB0">
            <w:r>
                <w:rPr>
                    <w:noProof/>
                </w:rPr>
                <w:drawing>
                    <wp:inline distT="T-mock-inline-p3-c0" distB="B-mock-inline-p3-c0" distL="L-mock-inline-p3-c0" distR="R-mock-inline-p3-c0" wp14:anchorId="03DD2422" wp14:editId="523B0CEC">
                        <wp:extent cx="2339163" cy="1494293"/>
                        <wp:effectExtent l="0" t="0" r="0" b="4445"/>
                        <wp:docPr id="2" name="图片 2"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="1" name="图片 1"/>
                                        <pic:cNvPicPr/>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="rId5" cstate="print">
                                            <a:extLst>
                                                <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                    <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                                </a:ext>
                                            </a:extLst>
                                        </a:blip>
                                        <a:stretch>
                                            <a:fillRect/>
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:off x="0" y="0"/>
                                            <a:ext cx="2378569" cy="1519466"/>
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst/>
                                        </a:prstGeom>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
            <w:r>
                <w:rPr>
                    <w:noProof/>
                </w:rPr>
                <w:drawing>
                    <wp:inline distT="T-mock-inline-p3-c1" distB="B-mock-inline-p3-c1" distL="L-mock-inline-p3-c1" distR="R-mock-inline-p3-c1" wp14:anchorId="6CAAB9D4" wp14:editId="6CA7D9C6">
                        <wp:extent cx="2339163" cy="1494293"/>
                        <wp:effectExtent l="0" t="0" r="0" b="4445"/>
                        <wp:docPr id="4" name="图片 4"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="1" name="图片 1"/>
                                        <pic:cNvPicPr/>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="rId5" cstate="print">
                                            <a:extLst>
                                                <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                    <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                                </a:ext>
                                            </a:extLst>
                                        </a:blip>
                                        <a:stretch>
                                            <a:fillRect/>
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:off x="0" y="0"/>
                                            <a:ext cx="2378569" cy="1519466"/>
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst/>
                                        </a:prstGeom>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                    </wp:inline>
                </w:drawing>
            </w:r>
        </w:p>
        <w:p w14:paraId="72A11E8B" w14:textId="0CEBAF2D" w:rsidR="00084CD5" w:rsidRDefault="00084CD5">
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
                <w:t xml:space="preserve">一行2个组合 </w:t>
            </w:r>
            <w:r>
                <w:t>inline</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="6C280000" w14:textId="64F61F86" w:rsidR="00EE4CB0" w:rsidRDefault="0056103D">
            <w:r>
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                    <w:noProof/>
                </w:rPr>
                <mc:AlternateContent>
                    <mc:Choice Requires="wpg">
                        <w:drawing>
                            <wp:inline distT="T-mock-inline-p5-c0" distB="B-mock-inline-p5-c0" distL="L-mock-inline-p5-c0" distR="R-mock-inline-p5-c0" wp14:anchorId="5843EF5F" wp14:editId="6D5EB296">
                                <wp:extent cx="4677868" cy="1494155"/>
                                <wp:effectExtent l="0" t="0" r="0" b="4445"/>
                                <wp:docPr id="7" name="组合 7"/>
                                <wp:cNvGraphicFramePr/>
                                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                    <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
                                        <wpg:wgp>
                                            <wpg:cNvGrpSpPr/>
                                            <wpg:grpSpPr>
                                                <a:xfrm>
                                                    <a:off x="0" y="0"/>
                                                    <a:ext cx="4677868" cy="1494155"/>
                                                    <a:chOff x="0" y="0"/>
                                                    <a:chExt cx="4677868" cy="1494155"/>
                                                </a:xfrm>
                                            </wpg:grpSpPr>
                                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                <pic:nvPicPr>
                                                    <pic:cNvPr id="6" name="图片 6"/>
                                                    <pic:cNvPicPr>
                                                        <a:picLocks noChangeAspect="1"/>
                                                    </pic:cNvPicPr>
                                                </pic:nvPicPr>
                                                <pic:blipFill>
                                                    <a:blip r:embed="rId6" cstate="print">
                                                        <a:extLst>
                                                            <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                                <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                                            </a:ext>
                                                        </a:extLst>
                                                    </a:blip>
                                                    <a:stretch>
                                                        <a:fillRect/>
                                                    </a:stretch>
                                                </pic:blipFill>
                                                <pic:spPr>
                                                    <a:xfrm>
                                                        <a:off x="2339163" y="0"/>
                                                        <a:ext cx="2338705" cy="1494155"/>
                                                    </a:xfrm>
                                                    <a:prstGeom prst="rect">
                                                        <a:avLst/>
                                                    </a:prstGeom>
                                                </pic:spPr>
                                            </pic:pic>
                                            <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                <pic:nvPicPr>
                                                    <pic:cNvPr id="5" name="图片 5"/>
                                                    <pic:cNvPicPr>
                                                        <a:picLocks noChangeAspect="1"/>
                                                    </pic:cNvPicPr>
                                                </pic:nvPicPr>
                                                <pic:blipFill>
                                                    <a:blip r:embed="rId6" cstate="print">
                                                        <a:extLst>
                                                            <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                                <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                                            </a:ext>
                                                        </a:extLst>
                                                    </a:blip>
                                                    <a:stretch>
                                                        <a:fillRect/>
                                                    </a:stretch>
                                                </pic:blipFill>
                                                <pic:spPr>
                                                    <a:xfrm>
                                                        <a:off x="0" y="0"/>
                                                        <a:ext cx="2338705" cy="1494155"/>
                                                    </a:xfrm>
                                                    <a:prstGeom prst="rect">
                                                        <a:avLst/>
                                                    </a:prstGeom>
                                                </pic:spPr>
                                            </pic:pic>
                                        </wpg:wgp>
                                    </a:graphicData>
                                </a:graphic>
                            </wp:inline>
                        </w:drawing>
                    </mc:Choice>
                    <mc:Fallback>
                        <w:pict>
                            <v:group w14:anchorId="27116046" id="组合 7" o:spid="_x0000_s1026" style="width:368.35pt;height:117.65pt;mso-position-horizontal-relative:char;mso-position-vertical-relative:line" coordsize="46778,14941" o:gfxdata="UEsDBBQABgAIAAAAIQCxgme2CgEAABMCAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbJSRwU7DMAyG&#xD;&#xA;70i8Q5QralN2QAit3YGOIyA0HiBK3DaicaI4lO3tSbpNgokh7Rjb3+8vyXK1tSObIJBxWPPbsuIM&#xD;&#xA;UDltsK/5++apuOeMokQtR4dQ8x0QXzXXV8vNzgOxRCPVfIjRPwhBagArqXQeMHU6F6yM6Rh64aX6&#xD;&#xA;kD2IRVXdCeUwAsYi5gzeLFvo5OcY2XqbynsTjz1nj/u5vKrmxmY+18WfRICRThDp/WiUjOluYkJ9&#xD;&#xA;4lUcnMpEzjM0GE83SfzMhtz57fRzwYF7SY8ZjAb2KkN8ljaZCx1IwMK1TpX/Z2RJS4XrOqOgbAOt&#xD;&#xA;Z+rodC5buy8MMF0a3ibsDaZjupi/tPkGAAD//wMAUEsDBBQABgAIAAAAIQA4/SH/1gAAAJQBAAAL&#xD;&#xA;AAAAX3JlbHMvLnJlbHOkkMFqwzAMhu+DvYPRfXGawxijTi+j0GvpHsDYimMaW0Yy2fr2M4PBMnrb&#xD;&#xA;Ub/Q94l/f/hMi1qRJVI2sOt6UJgd+ZiDgffL8ekFlFSbvV0oo4EbChzGx4f9GRdb25HMsYhqlCwG&#xD;&#xA;5lrLq9biZkxWOiqY22YiTra2kYMu1l1tQD30/bPm3wwYN0x18gb45AdQl1tp5j/sFB2T0FQ7R0nT&#xD;&#xA;NEV3j6o9feQzro1iOWA14Fm+Q8a1a8+Bvu/d/dMb2JY5uiPbhG/ktn4cqGU/er3pcvwCAAD//wMA&#xD;&#xA;UEsDBBQABgAIAAAAIQDcfKgKYwIAAFoHAAAOAAAAZHJzL2Uyb0RvYy54bWzsVVtqGzEU/S90D0L/&#xD;&#xA;8fhtZ7AdSt2YQmhNHwuQNZoZkdGDK/m1gtI1dC/dTek2eqWZuLEdcMlHodAPy9Lo6uqcc4+kyc1O&#xD;&#xA;VWQjwEmjp7TTalMiNDeZ1MWUfv50ezWmxHmmM1YZLaZ0Lxy9mb18MdnaVHRNaapMAMEk2qVbO6Wl&#xD;&#xA;9zZNEsdLoZhrGSs0TuYGFPM4hCLJgG0xu6qSbrs9TLYGMguGC+fw67yepLOYP88F9+/z3AlPqilF&#xD;&#xA;bD62ENtVaJPZhKUFMFtK3sBgz0ChmNS46SHVnHlG1iDPUinJwTiT+xY3KjF5LrmIHJBNp33CZgFm&#xD;&#xA;bSOXIt0W9iATSnui07PT8nebBdiPdgmoxNYWqEUcBS67HFT4R5RkFyXbHyQTO084fuwPR6PxEIvM&#xD;&#xA;ca7Tv+53BoNaVF6i8mfrePnmwsrkYePkCI6VPMVfowH2zjS47BVc5dcgaJNE/VEOxeB+ba+wXJZ5&#xD;&#xA;uZKV9PtoPSxMAKU3S8mXUA9QziUQmU3pkBLNFDr+x7fvP79+IcOgSogPIfUCFgjdGX7viDavS6YL&#xD;&#xA;8cpZtCwqGaKT4/A4PNptVUl7K6sqFCn0G15o7xN7PCFNbb254WsltK/PEogKKRrtSmkdJZAKtRLI&#xD;&#xA;Bd5mHawwnmOPhCxI7esaOw/......">
                                <v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t" path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
                                    <v:stroke joinstyle="miter"/>
                                    <v:formulas>
                                        <v:f eqn="if lineDrawn pixelLineWidth 0"/>
                                        <v:f eqn="sum @0 1 0"/>
                                        <v:f eqn="sum 0 0 @1"/>
                                        <v:f eqn="prod @2 1 2"/>
                                        <v:f eqn="prod @3 21600 pixelWidth"/>
                                        <v:f eqn="prod @3 21600 pixelHeight"/>
                                        <v:f eqn="sum @0 0 1"/>
                                        <v:f eqn="prod @6 1 2"/>
                                        <v:f eqn="prod @7 21600 pixelWidth"/>
                                        <v:f eqn="sum @8 21600 0"/>
                                        <v:f eqn="prod @7 21600 pixelHeight"/>
                                        <v:f eqn="sum @10 21600 0"/>
                                    </v:formulas>
                                    <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
                                    <o:lock v:ext="edit" aspectratio="t"/>
                                </v:shapetype>
                                <v:shape id="图片 6" o:spid="_x0000_s1027" type="#_x0000_t75" style="position:absolute;left:23391;width:23387;height:14941;visibility:visible;mso-wrap-style:square" o:gfxdata="UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH&#xD;&#xA;70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23&#xD;&#xA;jocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3&#xD;&#xA;VXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0&#xD;&#xA;r4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//&#xD;&#xA;AwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0&#xD;&#xA;X5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w&#xD;&#xA;Wl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR&#xD;&#xA;gPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A&#xD;&#xA;AAD//wMAUEsDBBQABgAIAAAAIQDK/ILfyAAAAN8AAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9Ba8JA&#xD;&#xA;FITvgv9heYXezKYegkZXaZVSLx7UFnp8zT6TtNm3aXY1q7++WxC8DAzDfMPMl8E04kydqy0reEpS&#xD;&#xA;EMSF1TWXCt4Pr6MJCOeRNTaWScGFHCwXw8Ecc2173tF570sRIexyVFB53+ZSuqIigy6xLXHMjrYz&#xD;&#xA;6KPtSqk77CPcNHKcppk0WHNcqLClVUXFz/5kFHxc377T7Oto/e7yErbhs/zdTHulHh/CehbleQbC&#xD;&#xA;U/D3xg2x0Qoy+P8Tv4Bc/AEAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAA&#xD;&#xA;AAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUB&#xD;&#xA;AAALAAAAAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQDK/ILfyAAAAN8A&#xD;&#xA;AAAPAAAAAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA/AIAAAAA&#xD;&#xA;">
                                    <v:imagedata r:id="rId7" o:title=""/>
                                </v:shape>
                                <v:shape id="图片 5" o:spid="_x0000_s1028" type="#_x0000_t75" style="position:absolute;width:23387;height:14941;visibility:visible;mso-wrap-style:square" o:gfxdata="UEsDBBQABgAIAAAAIQDb4fbL7gAAAIUBAAATAAAAW0NvbnRlbnRfVHlwZXNdLnhtbHyQz07DMAyH&#xD;&#xA;70i8Q+QralM4IITa7kDhCAiNB7ASt43WOlEcyvb2pNu4IODoPz9/n1xv9vOkForiPDdwXVagiI23&#xD;&#xA;jocG3rdPxR0oScgWJ8/UwIEENu3lRb09BBKV0ywNjCmFe63FjDSjlD4Q50nv44wpl3HQAc0OB9I3&#xD;&#xA;VXWrjedEnIq03oC27qjHjympx31un0wiTQLq4bS4shrAECZnMGVTvbD9QSnOhDInjzsyuiBXWQP0&#xD;&#xA;r4R18jfgnHvJr4nOknrFmJ5xzhraRtHWf3Kkpfz/yGo5S+H73hkquyhdjr3R8m2lj09svwAAAP//&#xD;&#xA;AwBQSwMEFAAGAAgAAAAhAFr0LFu/AAAAFQEAAAsAAABfcmVscy8ucmVsc2zPwWrDMAwG4Ptg72B0&#xD;&#xA;X5TuUMaI01uh19I+gLGVxCy2jGSy9e1nemrHjpL4P0nD4SetZiPRyNnCruvBUPYcYp4tXC/Htw8w&#xD;&#xA;Wl0ObuVMFm6kcBhfX4Yzra62kC6xqGlKVgtLreUTUf1CyWnHhXKbTCzJ1VbKjMX5LzcTvvf9HuXR&#xD;&#xA;gPHJNKdgQU5hB+ZyK23zHztFL6w81c5zQp6m6P9TMfB3PtPWFCczVQtB9N4U2rp2HOA44NMz4y8A&#xD;&#xA;AAD//wMAUEsDBBQABgAIAAAAIQA6LhyoyAAAAN8AAAAPAAAAZHJzL2Rvd25yZXYueG1sRI9BawIx&#xD;&#xA;FITvgv8hPKE3zVqotKtRbEX00oO2BY/PzXN3dfOy3UQ39tc3guBlYBjmG2YyC6YSF2pcaVnBcJCA&#xD;&#xA;IM6sLjlX8P217L+CcB5ZY2WZFFzJwWza7Uww1bblDV22PhcRwi5FBYX3dSqlywoy6Aa2Jo7ZwTYG&#xD;&#xA;fbRNLnWDbYSbSj4nyUgaLDkuFFjTR0HZaXs2Cn7+VsdktD9Yv7m+h8+wy3/Xb61ST72wGEeZj0F4&#xD;&#xA;Cv7RuCPWWsEL3P7ELyCn/wAAAP//AwBQSwECLQAUAAYACAAAACEA2+H2y+4AAACFAQAAEwAAAAAA&#xD;&#xA;AAAAAAAAAAAAAAAAW0NvbnRlbnRfVHlwZXNdLnhtbFBLAQItABQABgAIAAAAIQBa9CxbvwAAABUB&#xD;&#xA;AAALAAAAAAAAAAAAAAAAAB8BAABfcmVscy8ucmVsc1BLAQItABQABgAIAAAAIQA6LhyoyAAAAN8A&#xD;&#xA;AAAPAAAAAAAAAAAAAAAAAAcCAABkcnMvZG93bnJldi54bWxQSwUGAAAAAAMAAwC3AAAA/AIAAAAA&#xD;&#xA;">
                                    <v:imagedata r:id="rId7" o:title=""/>
                                </v:shape>
                                <w10:anchorlock/>
                            </v:group>
                        </w:pict>
                    </mc:Fallback>
                </mc:AlternateContent>
            </w:r>
        </w:p>
        <w:p w14:paraId="51243B69" w14:textId="039F28AD" w:rsidR="0056103D" w:rsidRDefault="00AF7765">
            <w:r>
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
                <w:t>一个 浮于上方</w:t>
            </w:r>
            <w:r w:rsidR="009A179B">
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
                <w:t xml:space="preserve"> 右侧对齐</w:t>
            </w:r>
            <w:r w:rsidR="0079003D">
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
                <w:t xml:space="preserve"> 左</w:t>
            </w:r>
            <w:r w:rsidR="0079003D">
                <w:t xml:space="preserve">11.32cm </w:t>
            </w:r>
            <w:r w:rsidR="0079003D">
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
                <w:t>顶</w:t>
            </w:r>
            <w:r w:rsidR="0079003D">
                <w:t>23.73cm</w:t>
            </w:r>
        </w:p>
        <w:p w14:paraId="7D425D3C" w14:textId="05800C1E" w:rsidR="00AF7765" w:rsidRDefault="0079003D">
            <w:pPr>
                <w:rPr>
                    <w:rFonts w:hint="eastAsia"/>
                </w:rPr>
            </w:pPr>
            <w:r>
                <w:rPr>
                    <w:noProof/>
                </w:rPr>
                <w:drawing>
                    <wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1" wp14:anchorId="3218CDF8" wp14:editId="091B2914">
                        <wp:simplePos x="0" y="0"/>
                        <wp:positionH relativeFrom="column">
                            <wp:posOffset>2935605</wp:posOffset>
                        </wp:positionH>
                        <wp:positionV relativeFrom="paragraph">
                            <wp:posOffset>97790</wp:posOffset>
                        </wp:positionV>
                        <wp:extent cx="2339163" cy="1494293"/>
                        <wp:effectExtent l="0" t="0" r="0" b="4445"/>
                        <wp:wrapNone/>
                        <wp:docPr id="8" name="图片 8"/>
                        <wp:cNvGraphicFramePr>
                            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                        </wp:cNvGraphicFramePr>
                        <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:nvPicPr>
                                        <pic:cNvPr id="1" name="图片 1"/>
                                        <pic:cNvPicPr/>
                                    </pic:nvPicPr>
                                    <pic:blipFill>
                                        <a:blip r:embed="rId6" cstate="print">
                                            <a:extLst>
                                                <a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                    <a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/>
                                                </a:ext>
                                            </a:extLst>
                                        </a:blip>
                                        <a:stretch>
                                            <a:fillRect/>
                                        </a:stretch>
                                    </pic:blipFill>
                                    <pic:spPr>
                                        <a:xfrm>
                                            <a:off x="0" y="0"/>
                                            <a:ext cx="2339163" cy="1494293"/>
                                        </a:xfrm>
                                        <a:prstGeom prst="rect">
                                            <a:avLst/>
                                        </a:prstGeom>
                                    </pic:spPr>
                                </pic:pic>
                            </a:graphicData>
                        </a:graphic>
                        <wp14:sizeRelH relativeFrom="page">
                            <wp14:pctWidth>0</wp14:pctWidth>
                        </wp14:sizeRelH>
                        <wp14:sizeRelV relativeFrom="page">
                            <wp14:pctHeight>0</wp14:pctHeight>
                        </wp14:sizeRelV>
                    </wp:anchor>
                </w:drawing>
            </w:r>
        </w:p>
        <w:sectPr w:rsidR="00AF7765">
            <w:pgSz w:w="11906" w:h="16838"/>
            <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" w:header="851" w:footer="992" w:gutter="0"/>
            <w:cols w:space="425"/>
            <w:docGrid w:type="lines" w:linePitch="312"/>
        </w:sectPr>
    </w:body>
</w:document>`

func TestUnmarshalDrawingStructure(t *testing.T) {
	doc := Document{
		XMLW:    XMLNS_W,
		XMLR:    XMLNS_R,
		XMLWP:   XMLNS_WP,
		XMLName: xml.Name{Space: XMLNS_W, Local: "document"}}
	err := xml.Unmarshal(StringToBytes(drawing_doc), &doc)
	if err != nil {
		t.Fatal(err)
	}
	if len(doc.Body.Paragraphs) != 8 {
		t.Fatalf("We expected %d paragraphs, we got %d", 8, len(doc.Body.Paragraphs))
	}
	for i, p := range doc.Body.Paragraphs {
		if len(p.Children) == 0 {
			t.Fatalf("We were not able to parse paragraph %d", i)
		}
		for j, child := range p.Children {
			if child.Link == nil && child.Properties == nil && child.Run == nil {
				t.Fatalf("There are Paragraph children with all fields nil")
			}
			if child.Run != nil && child.Run.Text == nil && child.Run.InstrText == "" && child.Run.Drawing == nil {
				t.Fatalf("We have a run with no text and drawing")
			}
			if child.Link != nil && child.Link.ID == "" {
				t.Fatalf("We have a link without ID")
			}
			if child.Run != nil && child.Run.Drawing != nil {
				t.Log("fild drawing at aragraph", i, ", child", j)
				if child.Run.Drawing.Inline != nil {
					tail := "-mock-inline-p" + string(rune('0'+i)) + "-c" + string(rune('0'+j))
					if "T"+tail != child.Run.Drawing.Inline.DistT {
						t.Fatal("expect", "T"+tail, "but got", child.Run.Drawing.Inline.DistT)
					}
					if "B"+tail != child.Run.Drawing.Inline.DistB {
						t.Fatal("expect", "B"+tail, "but got", child.Run.Drawing.Inline.DistB)
					}
					if "L"+tail != child.Run.Drawing.Inline.DistL {
						t.Fatal("expect", "L"+tail, "but got", child.Run.Drawing.Inline.DistL)
					}
					if "R"+tail != child.Run.Drawing.Inline.DistR {
						t.Fatal("expect", "R"+tail, "but got", child.Run.Drawing.Inline.DistR)
					}
				}
			}
		}
	}
}

func TestMarshalDrawingStructure(t *testing.T) {
	w := New()
	// add new paragraph
	para1 := w.AddParagraph()
	// add text
	para1.AddText("直接粘贴 inline")

	para2 := w.AddParagraph()
	para2.AddText("test font size and color").Size("44").Color("ff0000")
	para2.AddText("test font size and color").Size("44").Color("ff0000")
	para2.AddText("test font size and color").Size("44").Color("ff0000")

	nextPara := w.AddParagraph()
	nextPara.AddLink("google", `http://google.com`)

	f, err := os.Create("test.xml")
	if err != nil {
		t.Fatal(err)
	}
	defer f.Close()
	_, err = marshaller{data: w.Document}.WriteTo(f)
	if err != nil {
		t.Fatal(err)
	}
	_, err = f.Seek(0, io.SeekStart)
	if err != nil {
		t.Fatal(err)
	}
	w = New()
	err = xml.NewDecoder(f).Decode(&w.Document)
	if err != nil {
		t.Fatal(err)
	}
	f1, err := os.Create("test1.xml")
	if err != nil {
		t.Fatal(err)
	}
	defer f1.Close()
	_, err = marshaller{data: w.Document}.WriteTo(f1)
	if err != nil {
		t.Fatal(err)
	}
	t.Fail()
}
