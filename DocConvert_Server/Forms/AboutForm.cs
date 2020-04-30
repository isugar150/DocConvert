﻿using System;
using System.Reflection;
using System.Windows.Forms;

namespace DocConvert_Server.Forms
{
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
        }

        private void AboutForm_Load(object sender, EventArgs e)
        {
            label2.Text = string.Format("ver. {0}", Assembly.GetExecutingAssembly().GetName().Version.ToString());
            listView1.Items.Add(new ListViewItem(new string[] { "", "DocConvert", "Attribution-NonCommercial-NoDerivatives 4.0 International (CC BY-NC-ND 4.0)", "문서 변환 서버" }));
            listView1.Items.Add(new ListViewItem(new string[] { "", "vtortola.WebSockets", "MIT License", "웹 소켓 통신 라이브러리" }));
            listView1.Items.Add(new ListViewItem(new string[] { "", "websocket-sharp", "MIT License", "웹 소켓 클라이언트 라이브러리" }));
            listView1.Items.Add(new ListViewItem(new string[] { "", "pdfium", "Apache License", "PDF에서 이미지로 변환해주는 라이브러리" }));
            listView1.Items.Add(new ListViewItem(new string[] { "", "SharpZipLib", "MIT License", "압축 관련 라이브러리" }));
            listView1.Items.Add(new ListViewItem(new string[] { "", "FluentFTP", "MIT License", "FTP 접속 라이브러리" }));
            listView1.Items.Add(new ListViewItem(new string[] { "", "NLOG", "BSD 3-Clause Revised License", "로깅 관련 라이브러리" }));
            listView1.Items.Add(new ListViewItem(new string[] { "", "log4j", "Apache License", "로깅 관련 라이브러리" }));
            listView1.Items.Add(new ListViewItem(new string[] { "", "phantomjs", "BSD License", "웹 페이지 캡쳐 라이브러리" }));
        }

        #region 컴포넌트 이벤트
        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count == 0)
                return;
            ListView.SelectedListViewItemCollection itemColl = listView1.SelectedItems;
            foreach (ListViewItem item in itemColl)
            {
                for (int i = 0; i < listView1.Items.Count; i++)
                    listView1.Items[i].SubItems[0].Text = "";
                listView1.Items[item.Index].SubItems[0].Text = "*";
                textBox1.Text = getLicenseInfo(listView1.Items[item.Index].SubItems[1].Text);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Dispose();
        }
        #endregion

        private string getLicenseInfo(string license)
        {
            #region License Info
            string docConvert = "Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License\r\nBy exercising the Licensed Rights (defined below), You accept and agree to be bound by the terms and conditions of this Creative Commons Attribution-NonCommercial-NoDerivatives 4.0 International Public License (\"Public License\"). To the extent this Public License may be interpreted as a contract, You are granted the Licensed Rights in consideration of Your acceptance of these terms and conditions, and the Licensor grants You such rights in consideration of benefits the Licensor receives from making the Licensed Material available under these terms and conditions.\r\n\r\nSection 1 \u2013 Definitions.\r\n\r\nAdapted Material means material subject to Copyright and Similar Rights that is derived from or based upon the Licensed Material and in which the Licensed Material is translated, altered, arranged, transformed, or otherwise modified in a manner requiring permission under the Copyright and Similar Rights held by the Licensor. For purposes of this Public License, where the Licensed Material is a musical work, performance, or sound recording, Adapted Material is always produced where the Licensed Material is synched in timed relation with a moving image.\r\nCopyright and Similar Rights means copyright and/or similar rights closely related to copyright including, without limitation, performance, broadcast, sound recording, and Sui Generis Database Rights, without regard to how the rights are labeled or categorized. For purposes of this Public License, the rights specified in Section 2(b)(1)-(2) are not Copyright and Similar Rights.\r\nEffective Technological Measures means those measures that, in the absence of proper authority, may not be circumvented under laws fulfilling obligations under Article 11 of the WIPO Copyright Treaty adopted on December 20, 1996, and/or similar international agreements.\r\nExceptions and Limitations means fair use, fair dealing, and/or any other exception or limitation to Copyright and Similar Rights that applies to Your use of the Licensed Material.\r\nLicensed Material means the artistic or literary work, database, or other material to which the Licensor applied this Public License.\r\nLicensed Rights means the rights granted to You subject to the terms and conditions of this Public License, which are limited to all Copyright and Similar Rights that apply to Your use of the Licensed Material and that the Licensor has authority to license.\r\nLicensor means the individual(s) or entity(ies) granting rights under this Public License.\r\nNonCommercial means not primarily intended for or directed towards commercial advantage or monetary compensation. For purposes of this Public License, the exchange of the Licensed Material for other material subject to Copyright and Similar Rights by digital file-sharing or similar means is NonCommercial provided there is no payment of monetary compensation in connection with the exchange.\r\nShare means to provide material to the public by any means or process that requires permission under the Licensed Rights, such as reproduction, public display, public performance, distribution, dissemination, communication, or importation, and to make material available to the public including in ways that members of the public may access the material from a place and at a time individually chosen by them.\r\nSui Generis Database Rights means rights other than copyright resulting from Directive 96/9/EC of the European Parliament and of the Council of 11 March 1996 on the legal protection of databases, as amended and/or succeeded, as well as other essentially equivalent rights anywhere in the world.\r\nYou means the individual or entity exercising the Licensed Rights under this Public License. Your has a corresponding meaning.\r\nSection 2 \u2013 Scope.\r\n\r\nLicense grant.\r\nSubject to the terms and conditions of this Public License, the Licensor hereby grants You a worldwide, royalty-free, non-sublicensable, non-exclusive, irrevocable license to exercise the Licensed Rights in the Licensed Material to:\r\nreproduce and Share the Licensed Material, in whole or in part, for NonCommercial purposes only; and\r\nproduce and reproduce, but not Share, Adapted Material for NonCommercial purposes only.\r\nExceptions and Limitations. For the avoidance of doubt, where Exceptions and Limitations apply to Your use, this Public License does not apply, and You do not need to comply with its terms and conditions.\r\nTerm. The term of this Public License is specified in Section 6(a).\r\nMedia and formats; technical modifications allowed. The Licensor authorizes You to exercise the Licensed Rights in all media and formats whether now known or hereafter created, and to make technical modifications necessary to do so. The Licensor waives and/or agrees not to assert any right or authority to forbid You from making technical modifications necessary to exercise the Licensed Rights, including technical modifications necessary to circumvent Effective Technological Measures. For purposes of this Public License, simply making modifications authorized by this Section 2(a)(4) never produces Adapted Material.\r\nDownstream recipients.\r\nOffer from the Licensor \u2013 Licensed Material. Every recipient of the Licensed Material automatically receives an offer from the Licensor to exercise the Licensed Rights under the terms and conditions of this Public License.\r\nNo downstream restrictions. You may not offer or impose any additional or different terms or conditions on, or apply any Effective Technological Measures to, the Licensed Material if doing so restricts exercise of the Licensed Rights by any recipient of the Licensed Material.\r\nNo endorsement. Nothing in this Public License constitutes or may be construed as permission to assert or imply that You are, or that Your use of the Licensed Material is, connected with, or sponsored, endorsed, or granted official status by, the Licensor or others designated to receive attribution as provided in Section 3(a)(1)(A)(i).\r\nOther rights.\r\n\r\nMoral rights, such as the right of integrity, are not licensed under this Public License, nor are publicity, privacy, and/or other similar personality rights; however, to the extent possible, the Licensor waives and/or agrees not to assert any such rights held by the Licensor to the limited extent necessary to allow You to exercise the Licensed Rights, but not otherwise.\r\nPatent and trademark rights are not licensed under this Public License.\r\nTo the extent possible, the Licensor waives any right to collect royalties from You for the exercise of the Licensed Rights, whether directly or through a collecting society under any voluntary or waivable statutory or compulsory licensing scheme. In all other cases the Licensor expressly reserves any right to collect such royalties, including when the Licensed Material is used other than for NonCommercial purposes.\r\nSection 3 \u2013 License Conditions.\r\n\r\nYour exercise of the Licensed Rights is expressly made subject to the following conditions.\r\n\r\nAttribution.\r\n\r\nIf You Share the Licensed Material, You must:\r\n\r\nretain the following if it is supplied by the Licensor with the Licensed Material:\r\nidentification of the creator(s) of the Licensed Material and any others designated to receive attribution, in any reasonable manner requested by the Licensor (including by pseudonym if designated);\r\na copyright notice;\r\na notice that refers to this Public License;\r\na notice that refers to the disclaimer of warranties;\r\na URI or hyperlink to the Licensed Material to the extent reasonably practicable;\r\nindicate if You modified the Licensed Material and retain an indication of any previous modifications; and\r\nindicate the Licensed Material is licensed under this Public License, and include the text of, or the URI or hyperlink to, this Public License.\r\nFor the avoidance of doubt, You do not have permission under this Public License to Share Adapted Material.\r\nYou may satisfy the conditions in Section 3(a)(1) in any reasonable manner based on the medium, means, and context in which You Share the Licensed Material. For example, it may be reasonable to satisfy the conditions by providing a URI or hyperlink to a resource that includes the required information.\r\nIf requested by the Licensor, You must remove any of the information required by Section 3(a)(1)(A) to the extent reasonably practicable.\r\nSection 4 \u2013 Sui Generis Database Rights.\r\n\r\nWhere the Licensed Rights include Sui Generis Database Rights that apply to Your use of the Licensed Material:\r\n\r\nfor the avoidance of doubt, Section 2(a)(1) grants You the right to extract, reuse, reproduce, and Share all or a substantial portion of the contents of the database for NonCommercial purposes only and provided You do not Share Adapted Material;\r\nif You include all or a substantial portion of the database contents in a database in which You have Sui Generis Database Rights, then the database in which You have Sui Generis Database Rights (but not its individual contents) is Adapted Material; and\r\nYou must comply with the conditions in Section 3(a) if You Share all or a substantial portion of the contents of the database.\r\nFor the avoidance of doubt, this Section 4 supplements and does not replace Your obligations under this Public License where the Licensed Rights include other Copyright and Similar Rights.\r\nSection 5 \u2013 Disclaimer of Warranties and Limitation of Liability.\r\n\r\nUnless otherwise separately undertaken by the Licensor, to the extent possible, the Licensor offers the Licensed Material as-is and as-available, and makes no representations or warranties of any kind concerning the Licensed Material, whether express, implied, statutory, or other. This includes, without limitation, warranties of title, merchantability, fitness for a particular purpose, non-infringement, absence of latent or other defects, accuracy, or the presence or absence of errors, whether or not known or discoverable. Where disclaimers of warranties are not allowed in full or in part, this disclaimer may not apply to You.\r\nTo the extent possible, in no event will the Licensor be liable to You on any legal theory (including, without limitation, negligence) or otherwise for any direct, special, indirect, incidental, consequential, punitive, exemplary, or other losses, costs, expenses, or damages arising out of this Public License or use of the Licensed Material, even if the Licensor has been advised of the possibility of such losses, costs, expenses, or damages. Where a limitation of liability is not allowed in full or in part, this limitation may not apply to You.\r\nThe disclaimer of warranties and limitation of liability provided above shall be interpreted in a manner that, to the extent possible, most closely approximates an absolute disclaimer and waiver of all liability.\r\nSection 6 \u2013 Term and Termination.\r\n\r\nThis Public License applies for the term of the Copyright and Similar Rights licensed here. However, if You fail to comply with this Public License, then Your rights under this Public License terminate automatically.\r\nWhere Your right to use the Licensed Material has terminated under Section 6(a), it reinstates:\r\n\r\nautomatically as of the date the violation is cured, provided it is cured within 30 days of Your discovery of the violation; or\r\nupon express reinstatement by the Licensor.\r\nFor the avoidance of doubt, this Section 6(b) does not affect any right the Licensor may have to seek remedies for Your violations of this Public License.\r\nFor the avoidance of doubt, the Licensor may also offer the Licensed Material under separate terms or conditions or stop distributing the Licensed Material at any time; however, doing so will not terminate this Public License.\r\nSections 1, 5, 6, 7, and 8 survive termination of this Public License.\r\nSection 7 \u2013 Other Terms and Conditions.\r\n\r\nThe Licensor shall not be bound by any additional or different terms or conditions communicated by You unless expressly agreed.\r\nAny arrangements, understandings, or agreements regarding the Licensed Material not stated herein are separate from and independent of the terms and conditions of this Public License.\r\nSection 8 \u2013 Interpretation.\r\n\r\nFor the avoidance of doubt, this Public License does not, and shall not be interpreted to, reduce, limit, restrict, or impose conditions on any use of the Licensed Material that could lawfully be made without permission under this Public License.\r\nTo the extent possible, if any provision of this Public License is deemed unenforceable, it shall be automatically reformed to the minimum extent necessary to make it enforceable. If the provision cannot be reformed, it shall be severed from this Public License without affecting the enforceability of the remaining terms and conditions.\r\nNo term or condition of this Public License will be waived and no failure to comply consented to unless expressly agreed to by the Licensor.\r\nNothing in this Public License constitutes or may be interpreted as a limitation upon, or waiver of, any privileges and immunities that apply to the Licensor or You, including from the legal processes of any jurisdiction or authority.";
            string webSockets = "The MIT License (MIT)\r\nCopyright (c) 2014 vtortola\r\n\r\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\r\n\r\nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\r\n\r\nTHE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.";
            string webSocketClient = "The MIT License (MIT)\r\n\r\nCopyright (c) 2010-2020 sta.blockhead\r\n\r\nPermission is hereby granted, free of charge, to any person obtaining a copy\r\nof this software and associated documentation files (the \"Software\"), to deal\r\nin the Software without restriction, including without limitation the rights\r\nto use, copy, modify, merge, publish, distribute, sublicense, and/or sell\r\ncopies of the Software, and to permit persons to whom the Software is\r\nfurnished to do so, subject to the following conditions:\r\n\r\nThe above copyright notice and this permission notice shall be included in\r\nall copies or substantial portions of the Software.\r\n\r\nTHE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR\r\nIMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,\r\nFITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE\r\nAUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER\r\nLIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,\r\nOUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN\r\nTHE SOFTWARE.";
            string pdfium = "Copyright 2014 PDFium Authors. All rights reserved.\r\n\r\nRedistribution and use in source and binary forms, with or without\r\nmodification, are permitted provided that the following conditions are\r\nmet:\r\n\r\n   * Redistributions of source code must retain the above copyright\r\nnotice, this list of conditions and the following disclaimer.\r\n   * Redistributions in binary form must reproduce the above\r\ncopyright notice, this list of conditions and the following disclaimer\r\nin the documentation and/or other materials provided with the\r\ndistribution.\r\n   * Neither the name of Google Inc. nor the names of its\r\ncontributors may be used to endorse or promote products derived from\r\nthis software without specific prior written permission.\r\n\r\nTHIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS\r\n\"AS IS\" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT\r\nLIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR\r\nA PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT\r\nOWNER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,\r\nSPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT\r\nLIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE,\r\nDATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY\r\nTHEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT\r\n(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE\r\nOF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.";
            string fluentFTP = "Copyright (c) 2015 Robin Rodricks and FluentFTP Contributors\r\n\r\nPermission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the \"Software\"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:\r\n\r\nThe above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.\r\n\r\nTHE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.\r\n";
            string nLog = "Copyright (c) 2004-2020 Jaroslaw Kowalski <jaak@jkowalski.net>, Kim Christensen, Julian Verdurmen\r\n\r\nAll rights reserved.\r\n\r\nRedistribution and use in source and binary forms, with or without \r\nmodification, are permitted provided that the following conditions \r\nare met:\r\n\r\n* Redistributions of source code must retain the above copyright notice, \r\n  this list of conditions and the following disclaimer. \r\n\r\n* Redistributions in binary form must reproduce the above copyright notice,\r\n  this list of conditions and the following disclaimer in the documentation\r\n  and/or other materials provided with the distribution. \r\n\r\n* Neither the name of Jaroslaw Kowalski nor the names of its \r\n  contributors may be used to endorse or promote products derived from this\r\n  software without specific prior written permission. \r\n\r\nTHIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS \"AS IS\"\r\nAND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE \r\nIMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE \r\nARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE \r\nLIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR \r\nCONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF\r\nSUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS \r\nINTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN \r\nCONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) \r\nARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF \r\nTHE POSSIBILITY OF SUCH DAMAGE.";
            string log4j = "                                 Apache License\r\n                           Version 2.0, January 2004\r\n                        http://www.apache.org/licenses/\r\n\r\n   TERMS AND CONDITIONS FOR USE, REPRODUCTION, AND DISTRIBUTION\r\n\r\n   1. Definitions.\r\n\r\n      \"License\" shall mean the terms and conditions for use, reproduction,\r\n      and distribution as defined by Sections 1 through 9 of this document.\r\n\r\n      \"Licensor\" shall mean the copyright owner or entity authorized by\r\n      the copyright owner that is granting the License.\r\n\r\n      \"Legal Entity\" shall mean the union of the acting entity and all\r\n      other entities that control, are controlled by, or are under common\r\n      control with that entity. For the purposes of this definition,\r\n      \"control\" means (i) the power, direct or indirect, to cause the\r\n      direction or management of such entity, whether by contract or\r\n      otherwise, or (ii) ownership of fifty percent (50%) or more of the\r\n      outstanding shares, or (iii) beneficial ownership of such entity.\r\n\r\n      \"You\" (or \"Your\") shall mean an individual or Legal Entity\r\n      exercising permissions granted by this License.\r\n\r\n      \"Source\" form shall mean the preferred form for making modifications,\r\n      including but not limited to software source code, documentation\r\n      source, and configuration files.\r\n\r\n      \"Object\" form shall mean any form resulting from mechanical\r\n      transformation or translation of a Source form, including but\r\n      not limited to compiled object code, generated documentation,\r\n      and conversions to other media types.\r\n\r\n      \"Work\" shall mean the work of authorship, whether in Source or\r\n      Object form, made available under the License, as indicated by a\r\n      copyright notice that is included in or attached to the work\r\n      (an example is provided in the Appendix below).\r\n\r\n      \"Derivative Works\" shall mean any work, whether in Source or Object\r\n      form, that is based on (or derived from) the Work and for which the\r\n      editorial revisions, annotations, elaborations, or other modifications\r\n      represent, as a whole, an original work of authorship. For the purposes\r\n      of this License, Derivative Works shall not include works that remain\r\n      separable from, or merely link (or bind by name) to the interfaces of,\r\n      the Work and Derivative Works thereof.\r\n\r\n      \"Contribution\" shall mean any work of authorship, including\r\n      the original version of the Work and any modifications or additions\r\n      to that Work or Derivative Works thereof, that is intentionally\r\n      submitted to Licensor for inclusion in the Work by the copyright owner\r\n      or by an individual or Legal Entity authorized to submit on behalf of\r\n      the copyright owner. For the purposes of this definition, \"submitted\"\r\n      means any form of electronic, verbal, or written communication sent\r\n      to the Licensor or its representatives, including but not limited to\r\n      communication on electronic mailing lists, source code control systems,\r\n      and issue tracking systems that are managed by, or on behalf of, the\r\n      Licensor for the purpose of discussing and improving the Work, but\r\n      excluding communication that is conspicuously marked or otherwise\r\n      designated in writing by the copyright owner as \"Not a Contribution.\"\r\n\r\n      \"Contributor\" shall mean Licensor and any individual or Legal Entity\r\n      on behalf of whom a Contribution has been received by Licensor and\r\n      subsequently incorporated within the Work.\r\n\r\n   2. Grant of Copyright License. Subject to the terms and conditions of\r\n      this License, each Contributor hereby grants to You a perpetual,\r\n      worldwide, non-exclusive, no-charge, royalty-free, irrevocable\r\n      copyright license to reproduce, prepare Derivative Works of,\r\n      publicly display, publicly perform, sublicense, and distribute the\r\n      Work and such Derivative Works in Source or Object form.\r\n\r\n   3. Grant of Patent License. Subject to the terms and conditions of\r\n      this License, each Contributor hereby grants to You a perpetual,\r\n      worldwide, non-exclusive, no-charge, royalty-free, irrevocable\r\n      (except as stated in this section) patent license to make, have made,\r\n      use, offer to sell, sell, import, and otherwise transfer the Work,\r\n      where such license applies only to those patent claims licensable\r\n      by such Contributor that are necessarily infringed by their\r\n      Contribution(s) alone or by combination of their Contribution(s)\r\n      with the Work to which such Contribution(s) was submitted. If You\r\n      institute patent litigation against any entity (including a\r\n      cross-claim or counterclaim in a lawsuit) alleging that the Work\r\n      or a Contribution incorporated within the Work constitutes direct\r\n      or contributory patent infringement, then any patent licenses\r\n      granted to You under this License for that Work shall terminate\r\n      as of the date such litigation is filed.\r\n\r\n   4. Redistribution. You may reproduce and distribute copies of the\r\n      Work or Derivative Works thereof in any medium, with or without\r\n      modifications, and in Source or Object form, provided that You\r\n      meet the following conditions:\r\n\r\n      (a) You must give any other recipients of the Work or\r\n          Derivative Works a copy of this License; and\r\n\r\n      (b) You must cause any modified files to carry prominent notices\r\n          stating that You changed the files; and\r\n\r\n      (c) You must retain, in the Source form of any Derivative Works\r\n          that You distribute, all copyright, patent, trademark, and\r\n          attribution notices from the Source form of the Work,\r\n          excluding those notices that do not pertain to any part of\r\n          the Derivative Works; and\r\n\r\n      (d) If the Work includes a \"NOTICE\" text file as part of its\r\n          distribution, then any Derivative Works that You distribute must\r\n          include a readable copy of the attribution notices contained\r\n          within such NOTICE file, excluding those notices that do not\r\n          pertain to any part of the Derivative Works, in at least one\r\n          of the following places: within a NOTICE text file distributed\r\n          as part of the Derivative Works; within the Source form or\r\n          documentation, if provided along with the Derivative Works; or,\r\n          within a display generated by the Derivative Works, if and\r\n          wherever such third-party notices normally appear. The contents\r\n          of the NOTICE file are for informational purposes only and\r\n          do not modify the License. You may add Your own attribution\r\n          notices within Derivative Works that You distribute, alongside\r\n          or as an addendum to the NOTICE text from the Work, provided\r\n          that such additional attribution notices cannot be construed\r\n          as modifying the License.\r\n\r\n      You may add Your own copyright statement to Your modifications and\r\n      may provide additional or different license terms and conditions\r\n      for use, reproduction, or distribution of Your modifications, or\r\n      for any such Derivative Works as a whole, provided Your use,\r\n      reproduction, and distribution of the Work otherwise complies with\r\n      the conditions stated in this License.\r\n\r\n   5. Submission of Contributions. Unless You explicitly state otherwise,\r\n      any Contribution intentionally submitted for inclusion in the Work\r\n      by You to the Licensor shall be under the terms and conditions of\r\n      this License, without any additional terms or conditions.\r\n      Notwithstanding the above, nothing herein shall supersede or modify\r\n      the terms of any separate license agreement you may have executed\r\n      with Licensor regarding such Contributions.\r\n\r\n   6. Trademarks. This License does not grant permission to use the trade\r\n      names, trademarks, service marks, or product names of the Licensor,\r\n      except as required for reasonable and customary use in describing the\r\n      origin of the Work and reproducing the content of the NOTICE file.\r\n\r\n   7. Disclaimer of Warranty. Unless required by applicable law or\r\n      agreed to in writing, Licensor provides the Work (and each\r\n      Contributor provides its Contributions) on an \"AS IS\" BASIS,\r\n      WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or\r\n      implied, including, without limitation, any warranties or conditions\r\n      of TITLE, NON-INFRINGEMENT, MERCHANTABILITY, or FITNESS FOR A\r\n      PARTICULAR PURPOSE. You are solely responsible for determining the\r\n      appropriateness of using or redistributing the Work and assume any\r\n      risks associated with Your exercise of permissions under this License.\r\n\r\n   8. Limitation of Liability. In no event and under no legal theory,\r\n      whether in tort (including negligence), contract, or otherwise,\r\n      unless required by applicable law (such as deliberate and grossly\r\n      negligent acts) or agreed to in writing, shall any Contributor be\r\n      liable to You for damages, including any direct, indirect, special,\r\n      incidental, or consequential damages of any character arising as a\r\n      result of this License or out of the use or inability to use the\r\n      Work (including but not limited to damages for loss of goodwill,\r\n      work stoppage, computer failure or malfunction, or any and all\r\n      other commercial damages or losses), even if such Contributor\r\n      has been advised of the possibility of such damages.\r\n\r\n   9. Accepting Warranty or Additional Liability. While redistributing\r\n      the Work or Derivative Works thereof, You may choose to offer,\r\n      and charge a fee for, acceptance of support, warranty, indemnity,\r\n      or other liability obligations and/or rights consistent with this\r\n      License. However, in accepting such obligations, You may act only\r\n      on Your own behalf and on Your sole responsibility, not on behalf\r\n      of any other Contributor, and only if You agree to indemnify,\r\n      defend, and hold each Contributor harmless for any liability\r\n      incurred by, or claims asserted against, such Contributor by reason\r\n      of your accepting any such warranty or additional liability.\r\n\r\n   END OF TERMS AND CONDITIONS\r\n\r\n   APPENDIX: How to apply the Apache License to your work.\r\n\r\n      To apply the Apache License to your work, attach the following\r\n      boilerplate notice, with the fields enclosed by brackets \"[]\"\r\n      replaced with your own identifying information. (Don\'t include\r\n      the brackets!)  The text should be enclosed in the appropriate\r\n      comment syntax for the file format. We also recommend that a\r\n      file or class name and description of purpose be included on the\r\n      same \"printed page\" as the copyright notice for easier\r\n      identification within third-party archives.\r\n\r\n   Copyright 1999-2005 The Apache Software Foundation\r\n\r\n   Licensed under the Apache License, Version 2.0 (the \"License\");\r\n   you may not use this file except in compliance with the License.\r\n   You may obtain a copy of the License at\r\n\r\n       http://www.apache.org/licenses/LICENSE-2.0\r\n\r\n   Unless required by applicable law or agreed to in writing, software\r\n   distributed under the License is distributed on an \"AS IS\" BASIS,\r\n   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.\r\n   See the License for the specific language governing permissions and\r\n   limitations under the License.";
            string phantomjs = "Redistribution and use in source and binary forms, with or without\r\nmodification, are permitted provided that the following conditions are met:\r\n\r\n  * Redistributions of source code must retain the above copyright\r\n    notice, this list of conditions and the following disclaimer.\r\n  * Redistributions in binary form must reproduce the above copyright\r\n    notice, this list of conditions and the following disclaimer in the\r\n    documentation and/or other materials provided with the distribution.\r\n  * Neither the name of the <organization> nor the\r\n    names of its contributors may be used to endorse or promote products\r\n    derived from this software without specific prior written permission.\r\n\r\nTHIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS \"AS IS\"\r\nAND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE\r\nIMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE\r\nARE DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY\r\nDIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES\r\n(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;\r\nLOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND\r\nON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT\r\n(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF\r\nTHIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.";
            string sharpziplib = "Copyright \u00A9 2000-2018 SharpZipLib Contributors\r\n\r\nPermission is hereby granted, free of charge, to any person obtaining a copy of this\r\nsoftware and associated documentation files (the \"Software\"), to deal in the Software\r\nwithout restriction, including without limitation the rights to use, copy, modify, merge,\r\npublish, distribute, sublicense, and/or sell copies of the Software, and to permit persons\r\nto whom the Software is furnished to do so, subject to the following conditions:\r\n\r\nThe above copyright notice and this permission notice shall be included in all copies or\r\nsubstantial portions of the Software.\r\n\r\nTHE SOFTWARE IS PROVIDED \"AS IS\", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED,\r\nINCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR\r\nPURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE\r\nFOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR\r\nOTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER\r\nDEALINGS IN THE SOFTWARE.";
            #endregion

            if (license.Equals("DocConvert"))
                return docConvert;
            else if (license.Equals("vtortola.WebSockets"))
                return webSockets;
            else if (license.Equals("websocket-sharp"))
                return webSocketClient;
            else if (license.Equals("pdfium"))
                return pdfium;
            else if (license.Equals("FluentFTP"))
                return fluentFTP;
            else if (license.Equals("NLOG"))
                return nLog;
            else if (license.Equals("log4j"))
                return log4j;
            else if (license.Equals("phantomjs"))
                return phantomjs;
            else if (license.Equals("SharpZipLib"))
                return sharpziplib;
            else
                return null;
        }
    }
}
