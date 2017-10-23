using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RtfPipe;
using System.Xml;

namespace RtfPipe.Tests
{
  [TestClass]
  public class Conversion
  {
    [TestMethod]
    public void Conversion_Formatting()
    {
      const string rtf = @"{\rtf1\ansi\ansicpg1252\uc1\deff1{\fonttbl
{\f0\fswiss\fcharset0\fprq2 Arial;}
{\f1\fswiss\fcharset0\fprq2 Verdana;}
{\f2\froman\fcharset2\fprq2 Symbol;}}
{\colortbl;\red255\green0\blue0;\red0\green255\blue0;\red0\green0\blue255;}
{\stylesheet{\s0\itap0\nowidctlpar\f0\fs24 [Normal];}}
{\*\generator TX_RTF32 14.0.520.501;}
\sectd
\pard\itap0\nowidctlpar\plain\f1\fs36
{Hellou RTF Wörld\par
\f0\fs24\par
with some symbols: ""{\f2\cf1\cb0\chcbpat3\i abc}""\par
\par
and some \b bold\b0, \i italic\i0, \ul underlined\ul0  and \strike strikethrough\strike0  text\par
\par
some nested styles: {\b\cf1 bold}, {\i\cf2 italic}, {\ul\cf3 underlined}, normal\par
\par
some combined styles: {\b bold+\i italic+\ul underlined} vs. normal\par
\par
and further: {\b bold, {\i bold+italic}, \ul bold+underlined{\plain, normal}}\par
\par
different ways: [A] == [\'41] == [\u65A] == {\uc0[\u65]}\par
more unicode: Unicode: [\u915G] - ANSI-fallback: [G]\par
same but different: ""{\upr{[G]}{\*\ud{[\uc0\u915]}}}""\par
\par
something to ignore: {\*\unsupportedtag {\b should not appear}{\i this neither}}{\b visible}.
\par }
}";
      const string html = @"<!DOCTYPE html ><html><head><meta http-equiv=""content-type"" content=""text/html; charset=UTF-8"" /></head><body><p><span style=""font-family:Verdana;font-size:18pt"">Hellou RTF Wörld</span></p><p>&nbsp;</p><p><span style=""font-family:Arial;font-size:12pt"">with some symbols: ""</span><i><span style=""color:#ff0000;background-color:#0000ff;font-family:Symbol;font-size:12pt"">abc</span></i><span style=""font-family:Arial;font-size:12pt"">""</span></p><p>&nbsp;</p><p><span style=""font-family:Arial;font-size:12pt"">and some </span><b><span style=""font-family:Arial;font-size:12pt"">bold</span></b><span style=""font-family:Arial;font-size:12pt"">, </span><i><span style=""font-family:Arial;font-size:12pt"">italic</span></i><span style=""font-family:Arial;font-size:12pt"">, </span><u><span style=""font-family:Arial;font-size:12pt"">underlined</span></u><span style=""font-family:Arial;font-size:12pt""> and </span><s><span style=""font-family:Arial;font-size:12pt"">strikethrough</span></s><span style=""font-family:Arial;font-size:12pt""> text</span></p><p>&nbsp;</p><p><span style=""font-family:Arial;font-size:12pt"">some nested styles: </span><b><span style=""color:#ff0000;font-family:Arial;font-size:12pt"">bold</span></b><span style=""font-family:Arial;font-size:12pt"">, </span><i><span style=""color:#00ff00;font-family:Arial;font-size:12pt"">italic</span></i><span style=""font-family:Arial;font-size:12pt"">, </span><u><span style=""color:#0000ff;font-family:Arial;font-size:12pt"">underlined</span></u><span style=""font-family:Arial;font-size:12pt"">, normal</span></p><p>&nbsp;</p><p><span style=""font-family:Arial;font-size:12pt"">some combined styles: </span><b><span style=""font-family:Arial;font-size:12pt"">bold+</span></b><b><i><span style=""font-family:Arial;font-size:12pt"">italic+</span></i></b><b><i><u><span style=""font-family:Arial;font-size:12pt"">underlined</span></u></i></b><span style=""font-family:Arial;font-size:12pt""> vs. normal</span></p><p>&nbsp;</p><p><span style=""font-family:Arial;font-size:12pt"">and further: </span><b><span style=""font-family:Arial;font-size:12pt"">bold, </span></b><b><i><span style=""font-family:Arial;font-size:12pt"">bold+italic</span></i></b><b><span style=""font-family:Arial;font-size:12pt"">, </span></b><b><u><span style=""font-family:Arial;font-size:12pt"">bold+underlined</span></u></b><span style=""font-family:Arial;font-size:12pt"">, normal</span></p><p>&nbsp;</p><p><span style=""font-family:Arial;font-size:12pt"">different ways: [A] == [A] == [A] == [A]</span></p><p><span style=""font-family:Arial;font-size:12pt"">more unicode: Unicode: [Γ] - ANSI-fallback: [G]</span></p><p><span style=""font-family:Arial;font-size:12pt"">same but different: ""[Γ]""</span></p><p>&nbsp;</p><p><span style=""font-family:Arial;font-size:12pt"">something to ignore: </span><b><span style=""font-family:Arial;font-size:12pt"">visible</span></b><span style=""font-family:Arial;font-size:12pt"">.</span></p></body></html>";
      var output = Rtf.ToHtml(rtf);
      Assert.AreEqual(html, output);
    }

    [TestMethod]
    public void Conversion_AttachmentRender()
    {
      const string rtf = @"{\rtf1\adeflang1025\ansi\ansicpg1252\uc1\adeff31507\deff0\stshfdbch31506\stshfloch31506\stshfhich31506\stshfbi31507\deflang1033\deflangfe1033\themelang1033\themelangfe0\themelangcs0{\fonttbl{\f0\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman{\*\falt Times New Roman};}{\f34\fbidi \froman\fcharset1\fprq2{\*\panose 02040503050406030204}Cambria Math{\*\falt Calisto MT};}{\f39\fbidi \fswiss\fcharset0\fprq2{\*\panose 020f0502020204030204}Calibri{\*\falt Times New Roman};}{\f31500\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman{\*\falt Times New Roman};}{\f31501\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman{\*\falt Times New Roman};}{\f31502\fbidi \fswiss\fcharset0\fprq2{\*\panose 020f0302020204030204}Calibri Light{\*\falt Calibri};}{\f31503\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman{\*\falt Times New Roman};}{\f31504\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman{\*\falt Times New Roman};}{\f31505\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman{\*\falt Times New Roman};}{\f31506\fbidi \fswiss\fcharset0\fprq2{\*\panose 020f0502020204030204}Calibri{\*\falt Times New Roman};}{\f31507\fbidi \froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman{\*\falt Times New Roman};}{\f348\fbidi \froman\fcharset238\fprq2 Times New Roman CE{\*\falt Times New Roman};}{\f349\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr{\*\falt Times New Roman};}{\f351\fbidi \froman\fcharset161\fprq2 Times New Roman Greek{\*\falt Times New Roman};}{\f352\fbidi \froman\fcharset162\fprq2 Times New Roman Tur{\*\falt Times New Roman};}{\f353\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew){\*\falt Times New Roman};}{\f354\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic){\*\falt Times New Roman};}{\f355\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic{\*\falt Times New Roman};}{\f356\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese){\*\falt Times New Roman};}{\f688\fbidi \froman\fcharset238\fprq2 Cambria Math CE{\*\falt Calisto MT};}{\f689\fbidi \froman\fcharset204\fprq2 Cambria Math Cyr{\*\falt Calisto MT};}{\f691\fbidi \froman\fcharset161\fprq2 Cambria Math Greek{\*\falt Calisto MT};}{\f692\fbidi \froman\fcharset162\fprq2 Cambria Math Tur{\*\falt Calisto MT};}{\f695\fbidi \froman\fcharset186\fprq2 Cambria Math Baltic{\*\falt Calisto MT};}{\f696\fbidi \froman\fcharset163\fprq2 Cambria Math (Vietnamese){\*\falt Calisto MT};}{\f738\fbidi \fswiss\fcharset238\fprq2 Calibri CE{\*\falt Times New Roman};}{\f739\fbidi \fswiss\fcharset204\fprq2 Calibri Cyr{\*\falt Times New Roman};}{\f741\fbidi \fswiss\fcharset161\fprq2 Calibri Greek{\*\falt Times New Roman};}{\f742\fbidi \fswiss\fcharset162\fprq2 Calibri Tur{\*\falt Times New Roman};}{\f745\fbidi \fswiss\fcharset186\fprq2 Calibri Baltic{\*\falt Times New Roman};}{\f746\fbidi \fswiss\fcharset163\fprq2 Calibri (Vietnamese){\*\falt Times New Roman};}{\f31508\fbidi \froman\fcharset238\fprq2 Times New Roman CE{\*\falt Times New Roman};}{\f31509\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr{\*\falt Times New Roman};}{\f31511\fbidi \froman\fcharset161\fprq2 Times New Roman Greek{\*\falt Times New Roman};}{\f31512\fbidi \froman\fcharset162\fprq2 Times New Roman Tur{\*\falt Times New Roman};}{\f31513\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew){\*\falt Times New Roman};}{\f31514\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic){\*\falt Times New Roman};}{\f31515\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic{\*\falt Times New Roman};}{\f31516\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese){\*\falt Times New Roman};}{\f31518\fbidi \froman\fcharset238\fprq2 Times New Roman CE{\*\falt Times New Roman};}{\f31519\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr{\*\falt Times New Roman};}{\f31521\fbidi \froman\fcharset161\fprq2 Times New Roman Greek{\*\falt Times New Roman};}{\f31522\fbidi \froman\fcharset162\fprq2 Times New Roman Tur{\*\falt Times New Roman};}{\f31523\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew){\*\falt Times New Roman};}{\f31524\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic){\*\falt Times New Roman};}{\f31525\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic{\*\falt Times New Roman};}{\f31526\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese){\*\falt Times New Roman};}{\f31528\fbidi \fswiss\fcharset238\fprq2 Calibri Light CE{\*\falt Calibri};}{\f31529\fbidi \fswiss\fcharset204\fprq2 Calibri Light Cyr{\*\falt Calibri};}{\f31531\fbidi \fswiss\fcharset161\fprq2 Calibri Light Greek{\*\falt Calibri};}{\f31532\fbidi \fswiss\fcharset162\fprq2 Calibri Light Tur{\*\falt Calibri};}{\f31535\fbidi \fswiss\fcharset186\fprq2 Calibri Light Baltic{\*\falt Calibri};}{\f31536\fbidi \fswiss\fcharset163\fprq2 Calibri Light (Vietnamese){\*\falt Calibri};}{\f31538\fbidi \froman\fcharset238\fprq2 Times New Roman CE{\*\falt Times New Roman};}{\f31539\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr{\*\falt Times New Roman};}{\f31541\fbidi \froman\fcharset161\fprq2 Times New Roman Greek{\*\falt Times New Roman};}{\f31542\fbidi \froman\fcharset162\fprq2 Times New Roman Tur{\*\falt Times New Roman};}{\f31543\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew){\*\falt Times New Roman};}{\f31544\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic){\*\falt Times New Roman};}{\f31545\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic{\*\falt Times New Roman};}{\f31546\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese){\*\falt Times New Roman};}{\f31548\fbidi \froman\fcharset238\fprq2 Times New Roman CE{\*\falt Times New Roman};}{\f31549\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr{\*\falt Times New Roman};}{\f31551\fbidi \froman\fcharset161\fprq2 Times New Roman Greek{\*\falt Times New Roman};}{\f31552\fbidi \froman\fcharset162\fprq2 Times New Roman Tur{\*\falt Times New Roman};}{\f31553\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew){\*\falt Times New Roman};}{\f31554\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic){\*\falt Times New Roman};}{\f31555\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic{\*\falt Times New Roman};}{\f31556\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese){\*\falt Times New Roman};}{\f31558\fbidi \froman\fcharset238\fprq2 Times New Roman CE{\*\falt Times New Roman};}{\f31559\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr{\*\falt Times New Roman};}{\f31561\fbidi \froman\fcharset161\fprq2 Times New Roman Greek{\*\falt Times New Roman};}{\f31562\fbidi \froman\fcharset162\fprq2 Times New Roman Tur{\*\falt Times New Roman};}{\f31563\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew){\*\falt Times New Roman};}{\f31564\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic){\*\falt Times New Roman};}{\f31565\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic{\*\falt Times New Roman};}{\f31566\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese){\*\falt Times New Roman};}{\f31568\fbidi \fswiss\fcharset238\fprq2 Calibri CE{\*\falt Times New Roman};}{\f31569\fbidi \fswiss\fcharset204\fprq2 Calibri Cyr{\*\falt Times New Roman};}{\f31571\fbidi \fswiss\fcharset161\fprq2 Calibri Greek{\*\falt Times New Roman};}{\f31572\fbidi \fswiss\fcharset162\fprq2 Calibri Tur{\*\falt Times New Roman};}{\f31575\fbidi \fswiss\fcharset186\fprq2 Calibri Baltic{\*\falt Times New Roman};}{\f31576\fbidi \fswiss\fcharset163\fprq2 Calibri (Vietnamese){\*\falt Times New Roman};}{\f31578\fbidi \froman\fcharset238\fprq2 Times New Roman CE{\*\falt Times New Roman};}{\f31579\fbidi \froman\fcharset204\fprq2 Times New Roman Cyr{\*\falt Times New Roman};}{\f31581\fbidi \froman\fcharset161\fprq2 Times New Roman Greek{\*\falt Times New Roman};}{\f31582\fbidi \froman\fcharset162\fprq2 Times New Roman Tur{\*\falt Times New Roman};}{\f31583\fbidi \froman\fcharset177\fprq2 Times New Roman (Hebrew){\*\falt Times New Roman};}{\f31584\fbidi \froman\fcharset178\fprq2 Times New Roman (Arabic){\*\falt Times New Roman};}{\f31585\fbidi \froman\fcharset186\fprq2 Times New Roman Baltic{\*\falt Times New Roman};}{\f31586\fbidi \froman\fcharset163\fprq2 Times New Roman (Vietnamese){\*\falt Times New Roman};}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;\red5\green99\blue193;\red149\green79\blue114;}{\*\defchp \f31506\fs22 }{\*\defpap \ql \li0\ri0\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 }\noqfpromote {\stylesheet{\ql \li0\ri0\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \rtlch\fcs1 \af31507\afs22\alang1025 \ltrch\fcs0 \f31506\fs22\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 \snext0 \sqformat \spriority0 Normal;}{\*\cs10 \additive \ssemihidden \sunhideused \spriority1 Default Paragraph Font;}{\*\ts11\tsrowd\trftsWidthB3\trpaddl108\trpaddr108\trpaddfl3\trpaddft3\trpaddfb3\trpaddfr3\tblind0\tblindtype3\tsvertalt\tsbrdrt\tsbrdrl\tsbrdrb\tsbrdrr\tsbrdrdgl\tsbrdrdgr\tsbrdrh\tsbrdrv \ql \li0\ri0\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \rtlch\fcs1 \af31507\afs22\alang1025 \ltrch\fcs0 \f31506\fs22\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 \snext11 \ssemihidden \sunhideused Normal Table;}{\*\cs15 \additive \rtlch\fcs1 \af0 \ltrch\fcs0 \ul\cf17 \sbasedon10 \ssemihidden \sunhideused \styrsid16650015 Hyperlink;}{\*\cs16 \additive \rtlch\fcs1 \af0 \ltrch\fcs0 \ul\cf18 \sbasedon10 \ssemihidden \sunhideused \styrsid16650015 FollowedHyperlink;}{\*\cs17 \additive \rtlch\fcs1 \af31507\afs22 \ltrch\fcs0 \f31506\fs22\cf0 \sbasedon10 \ssemihidden \spriority0 \spersonal \scompose \styrsid16650015 EmailStyle17;}}{\*\revtbl {Unknown;}}{\*\rsidtbl \rsid5455638\rsid7608263\rsid16650015}{\mmathPr\mmathFont34\mbrkBin0\mbrkBinSub0\msmallFrac0\mdispDef1\mlMargin0\mrMargin0\mdefJc1\mwrapIndent1440\mintLim0\mnaryLim1}{\*\xmlnstbl {\xmlns1 http://schemas.microsoft.com/office/word/2003/wordml}}\paperw12240\paperh15840\margl1440\margr1440\margt1440\margb1440\gutter0\ltrsect \widowctrl\ftnbj\aenddoc\trackmoves0\trackformatting1\donotembedsysfont1\relyonvml0\donotembedlingdata0\grfdocevents0\validatexml1\showplaceholdtext0\ignoremixedcontent0\saveinvalidxml0\showxmlerrors1\noxlattoyen\expshrtn\noultrlspc\dntblnsbdb\nospaceforul\formshade\horzdoc\dgmargin\dghspace180\dgvspace180\dghorigin150\dgvorigin0\dghshow1\dgvshow1\jexpand\viewkind5\viewscale100\pgbrdrhead\pgbrdrfoot\splytwnine\ftnlytwnine\htmautsp\nolnhtadjtbl\useltbaln\alntblind\lytcalctblwd\lyttblrtgr\lnbrkrule\nobrkwrptbl\snaptogridincell\allowfieldendsel\wrppunct\asianbrkrule\rsidroot7608263\newtblstyruls\nogrowautofit\usenormstyforlist\noindnmbrts\felnbrelev\nocxsptable\indrlsweleven\noafcnsttbl\afelev\utinl\hwelev\spltpgpar\notcvasp\notbrkcnstfrctbl\notvatxbx\krnprsnet\cachedcolbal \nouicompat \fet0{\*\wgrffmtfilter 2450}\nofeaturethrottle1\ilfomacatclnup0\ltrpar \sectd \ltrsect\linex0\endnhere\sectdefaultcl\sftnbj {\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang {\pntxta )}}{\*\pnseclvl5\pndec\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}\pard\plain \ltrpar\ql \li0\ri0\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0\pararsid16650015 \rtlch\fcs1 \af31507\afs22\alang1025 \ltrch\fcs0 \f31506\fs22\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 {\rtlch\fcs1 \af31507 \ltrch\fcs0 \cf0\insrsid16650015 Testing with an inline image
\par 
\par }{\rtlch\fcs1 \af31507 \ltrch\fcs0 \cf0\lang1024\langfe1024\noproof\insrsid7608263\charrsid14577418 \objattph  }{\rtlch\fcs1 \af31507 \ltrch\fcs0 \cf0\insrsid16650015 
\par }\sectd \ltrsect\linex0\endnhere\sectdefaultcl\sftnbj {\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang {\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang {\pntxta )}}{\*\pnseclvl5\pndec\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang {\pntxtb (}{\pntxta )}}\pard\plain \ltrpar\ql \li0\ri0\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0\pararsid16650015 \rtlch\fcs1 \af31507\afs22\alang1025 \ltrch\fcs0 \f31506\fs22\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 {\rtlch\fcs1 \af31507 \ltrch\fcs0 \cf0\insrsid16650015 
\par And a file}{\rtlch\fcs1 \af31507 \ltrch\fcs0 \cf0\insrsid16650015 
\par 
\par }{\pard\plain \ltrpar\ql \li0\ri0\widctlpar\wrapdefault\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0\pararsid16650015 \rtlch\fcs1 \af31507\afs22\alang1025 \ltrch\fcs0 \f31506\fs22\lang1033\langfe1033\cgrid\langnp1033\langfenp1033\insrsid16650015 {{\objattph  {\rtlch\fcs1 \af31507 \ltrch\fcs0 \cf0\insrsid16650015 }}}}\sectd \ltrsect\linex0\endnhere\sectdefaultcl\sftnbj {\rtlch\fcs1 \af31507 \ltrch\fcs0 \cf0\insrsid16650015\charrsid16650015 
\par }}";
      const string html = "<!DOCTYPE html ><html><head><meta http-equiv=\"content-type\" content=\"text/html; charset=UTF-8\" /></head><body><p><span style=\"font-family:Calibri;font-size:11pt\">Testing with an inline image</span></p><p>&nbsp;</p><p><div data-index=\"0\" /><span style=\"font-family:Calibri;font-size:11pt\"> </span></p><p>&nbsp;</p><p><span style=\"font-family:Calibri;font-size:11pt\">And a file</span></p><p>&nbsp;</p><p><div data-index=\"1\" /><span style=\"font-family:Calibri;font-size:11pt\"> </span></p></body></html>";

      var settings = new RtfHtmlWriterSettings();
      settings.ObjectVisitor = new Visitor();
      var output = Rtf.ToHtml(rtf, settings);

      Assert.AreEqual(html, output);
    }

    [TestMethod]
    public void Conversion_ImageSize()
    {
      const string rtf = @"{\rtf1\ansi\ansicpg1251\deff0\nouicompat\deflang1049{\fonttbl{\f0\fnil\fcharset0 Calibri;}}
{\*\generator Riched20 10.0.14393}\viewkind4\uc1 
\pard\sa200\sl240\slmult1\f0\fs22\lang9{\pict{\*\picprop}\wmetafile8\picw1323\pich265\picwgoal750\pichgoal150 
010009000003f60000000000cd00000000000400000003010800050000000b0200000000050000
000c020a003200030000001e0004000000070104000400000007010400cd000000410b2000cc00
0a003200000000000a0032000000000028000000320000000a0000000100040000000000000000
000000000000000000000000000000000000000000ffffff003300ff000033ff00000000000000
000000000000000000000000000000000000000000000000000000000000000000000000000000
000000222222222222222222222222222222222222222222222222220202022222222222222222
222222222222222222222222222222222202020222222222222222222222222222222222222222
222222222222020202222222222222222222222222222222222222222222222222220202022222
222222222222222222222222222222222222222222222202020222222222222222222222222222
222222222222222222222222020202222222222222222222222222222222222222222222222222
220202022222222222222222222222222222222222222222222222222202020222222222222222
222222222222222222222222222222222222020202222222222222222222222222222222222222
22222222222232020202040000002701ffff030000000000
}\par
\par
\par
{\pict{\*\picprop}\wmetafile8\picw1323\pich265\picwgoal750\pichgoal150 
0100090000037600000000004d00000000000400000003010800050000000b0200000000050000
000c020a003200030000001e00040000000701040004000000070104004d000000410b2000cc00
0a003200000000000a0032000000000028000000320000000a0000000100010000000000000000
000000000000000000000000000000000000000000ffffff00ffffffffffffc001ffffffffffff
c001ffffffffffffc001ffffffffffffc001ffffffffffffc001ffffffffffffc001ffffffffff
ffc001ffffffffffffc001ffffffffffffc001ffffffffffffc001040000002701ffff03000000
0000
}\par

\pard\sa200\sl276\slmult1\par
}";
      const string html = "<!DOCTYPE html ><html><head><meta http-equiv=\"co" +
        "ntent-type\" content=\"text/html; charset=UTF-8\" /></head><body><p" +
        "><img width=\"50\" height=\"10\" src=\"data:windows/metafile;base64" +
        ",AQAJAAAD9gAAAAAAzQAAAAAABAAAAAMBCAAFAAAACwIAAAAABQAAAAwCCgAyAAMAAA" +
        "AeAAQAAAAHAQQABAAAAAcBBADNAAAAQQsgAMwACgAyAAAAAAAKADIAAAAAACgAAAAyA" +
        "AAACgAAAAEABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA////ADMA/wAAM/8A" +
        "AAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAIiI" +
        "iIiIiIiIiIiIiIiIiIiIiIiIiIiIiIgICAiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIi" +
        "ICAgIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiAgICIiIiIiIiIiIiIiIiIiIiIiIiI" +
        "iIiIiIiIgICAiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiICAgIiIiIiIiIiIiIiIiIi" +
        "IiIiIiIiIiIiIiIiAgICIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIgICAiIiIiIiIiI" +
        "iIiIiIiIiIiIiIiIiIiIiIiICAgIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiIiAgICIi" +
        "IiIiIiIiIiIiIiIiIiIiIiIiIiIiIiMgICAgQAAAAnAf//AwAAAAAA\" /></p><p>&" +
        "nbsp;</p><p>&nbsp;</p><p><img width=\"50\" height=\"10\" src=\"data" +
        ":windows/metafile;base64,AQAJAAADdgAAAAAATQAAAAAABAAAAAMBCAAFAAAACw" +
        "IAAAAABQAAAAwCCgAyAAMAAAAeAAQAAAAHAQQABAAAAAcBBABNAAAAQQsgAMwACgAyA" +
        "AAAAAAKADIAAAAAACgAAAAyAAAACgAAAAEAAQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAA" +
        "AAAAAAAA////AP///////8AB////////wAH////////AAf///////8AB////////wAH" +
        "////////AAf///////8AB////////wAH////////AAf///////8ABBAAAACcB//8DAA" +
        "AAAAA=\" /></p><p>&nbsp;</p></body></html>";

      var settings = new RtfHtmlWriterSettings();
      settings.ObjectVisitor = new Visitor();
      var output = Rtf.ToHtml(rtf, settings);

      Assert.AreEqual(html, output);
    }

    private class Visitor : DataUriImageVisitor
    {
      public override void RenderObject(int index, XmlWriter writer)
      {
        writer.WriteStartElement("div");
        writer.WriteAttributeString("data-index", index.ToString());
        writer.WriteEndElement();
      }
    }
  }
}
