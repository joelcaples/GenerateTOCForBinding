<!--
  Licensed to the Apache Software Foundation (ASF) under one or more
  contributor license agreements.  See the NOTICE file distributed with
  this work for additional information regarding copyright ownership.
  The ASF licenses this file to You under the Apache License, Version 2.0
  (the "License"); you may not use this file except in compliance with
  the License.  You may obtain a copy of the License at

       http://www.apache.org/licenses/LICENSE-2.0

  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.
-->
<!-- $Id: xml2pdf.xsl 426576 2006-07-28 15:44:37Z jeremias $ -->
<xsl:stylesheet
     xmlns:xsl="http://www.w3.org/1999/XSL/Transform" version="1.0"
     xmlns:fo="http://www.w3.org/1999/XSL/Format">

  <xsl:variable name="MAX_ROWS">5</xsl:variable>
  <xsl:variable name="MAX_COLS">2</xsl:variable>

  <xsl:template match ="root">

    <fo:root xmlns:fo="http://www.w3.org/1999/XSL/Format">
      <fo:layout-master-set>
        <fo:simple-page-master master-name="first"
            margin-right="0in"
            margin-left="0in"
            margin-bottom="0in"
            margin-top="0in"
            page-width="6.625in"
            page-height="10.25in">
          <fo:region-body 
            background-color="#191970" 
            color="#FFFFFF"></fo:region-body>
        </fo:simple-page-master>

        <fo:simple-page-master master-name="rest"
            margin-right="0in"
            margin-left="0in"
            margin-bottom="0in"
            margin-top="0in"
            page-width="6.625in"
            page-height="10.25in">
          <fo:region-body 
            background-color="#191970" 
            color="#FFFFFF"></fo:region-body>
        </fo:simple-page-master>

        <fo:page-sequence-master master-name="psmA">
          <fo:repeatable-page-master-alternatives>
            <fo:conditional-page-master-reference master-reference="first" page-position="first" />
            <fo:conditional-page-master-reference master-reference="rest" page-position="rest" />
            <!-- recommended fallback procedure -->
            <fo:conditional-page-master-reference master-reference="rest" />
          </fo:repeatable-page-master-alternatives>
        </fo:page-sequence-master>
      </fo:layout-master-set>

      <fo:page-sequence master-reference="psmA">
        <fo:flow flow-name="xsl-region-body">

          <xsl:for-each select="/root/volumes/volume">

            <xsl:call-template name="volume_page" />

            <xsl:variable name="volume_idx">
              <xsl:value-of select="position()"/>
            </xsl:variable>

            <xsl:for-each select="./issue">
              <xsl:variable name="issue_idx_1">
                <xsl:value-of select="position()"/>
              </xsl:variable>
              <!--$issue_idx_1:<xsl:value-of select="$issue_idx_1"/><br/-->
              <xsl:if test="(position() - 1) mod ($MAX_ROWS * $MAX_COLS) = 0">
                
                <!--table width="1988px" height="3075px" border="1" cellspacing="0"-->
                  <!--tr-->
                    <!--td class="volume_title" align="center" valign="top"-->

                      <xsl:for-each select="../issue">
                        <xsl:variable name="issue_idx_2">
                          <xsl:value-of select="position()"/>
                        </xsl:variable>
                        <!--$issue_idx_2:<xsl:value-of select="$issue_idx_2"/><br/-->
                        <!--($issue_idx_2 &gt;= $issue_idx_1) and ($issue_idx_2 &lt;= $issue_idx_1 + ($MAX_ROWS * $MAX_COLS)):<xsl:value-of select="($issue_idx_2 &gt;= $issue_idx_1) and ($issue_idx_2 &lt;= $issue_idx_1 + ($MAX_ROWS * $MAX_COLS))"/><br/-->
                        <xsl:if test="($issue_idx_2 &gt;= $issue_idx_1) and ($issue_idx_2 &lt; $issue_idx_1 + $MAX_ROWS)">
                          <xsl:choose>

                            <xsl:when test="ceiling((position() div $MAX_ROWS) mod 2) = 1">

                              <xsl:call-template name="issue_group">
                                <xsl:with-param name="pos">
                                  <xsl:value-of select="position()"/>
                                </xsl:with-param>
                              </xsl:call-template>

                              <xsl:choose>
                                <xsl:when test="position() = count(../issue)">
                                  <!--*** FOOTER2 ***-->
                                  <!--p class="footer" /-->
                                  <fo:block break-after='page'/>

                                </xsl:when>
                                <xsl:when test="position() mod $MAX_ROWS*$MAX_COLS = 0">
                                  <!--*** FOOTER1 ***-->
                                  <!--p class="footer" /-->
                                  <fo:block break-after='page'/>

                                </xsl:when>
                              </xsl:choose>

                            </xsl:when>

                          </xsl:choose>
                        </xsl:if>
                      </xsl:for-each>
                    <!--/td-->
                  <!--/tr-->
                <!--/table-->
              </xsl:if>
            </xsl:for-each>

          </xsl:for-each>
<!--
          <fo:block-container border-color="black" border-style="solid" border-width=".5mm" height="3cm" width="8.95cm" top="3.67cm" left="0cm" padding=".6mm" position="absolute">
            <fo:block text-align="start" space-after.optimum="3pt" line-height="14pt" font-family="sans-serif" font-size="12pt">
            </fo:block>
          </fo:block-container>

          <fo:block-container border-color="black" border-style="solid" border-width=".5mm" height="3cm" width="8.95cm" top="3.67cm" left="10cm" padding=".6mm" position="absolute">
            <fo:block text-align="start" space-after.optimum="3pt" line-height="14pt" font-family="sans-serif" font-size="12pt">
            </fo:block>
          </fo:block-container>
-->

        </fo:flow>
      </fo:page-sequence>
    </fo:root>
  </xsl:template>

  <xsl:template name="volume_page">

    <!--table width="1988px" height="3075px" border="1" cellspacing="0"-->
      <!--tr-->
        <!--td class="volume_title" align="center" valign="top"-->

    <fo:block space-after="2in" />

    <fo:block margin="1in">
      <fo:external-graphic height="2.25in" content-height="scale-to-fit" scaling="uniform" content-width="100%" src="url('./FF_LOGO_mb.gif')" />
    </fo:block>

    <!--br /-->
      <fo:block text-align="start" space-after.optimum="3pt" line-height="14pt" font-family="sans-serif" font-size="12pt" color="#FFFFFF" margin=".5in">

    <xsl:choose>
            <xsl:when test="@volume_code='FFBYRNE_1'">

              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                 <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="center">John Byrne's Fantastic Four Volume 1</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
              
              <fo:block space-after=".25in" />

              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                  <fo:table-row>
                    <fo:table-cell><fo:block text-align="right">Fantastic Four</fo:block></fo:table-cell>
                    <fo:table-cell><fo:block> #215-218</fo:block></fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell><fo:block> </fo:block></fo:table-cell>
                    <fo:table-cell><fo:block> #220-221</fo:block></fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell><fo:block> </fo:block></fo:table-cell>
                    <fo:table-cell><fo:block> #232-239</fo:block></fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell><fo:block text-align="right">Marvel Team Up</fo:block></fo:table-cell>
                    <fo:table-cell><fo:block> #61-62</fo:block></fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell><fo:block text-align="right">Marvel Two-In-One</fo:block></fo:table-cell>
                    <fo:table-cell><fo:block> #50</fo:block></fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
            </xsl:when>
            <xsl:when test="@volume_code='FFBYRNE_2'">

              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                 <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="center">John Byrne's Fantastic Four Volume 2</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
              
              <fo:block space-after=".25in" />


              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                  <fo:table-row>
                    <fo:table-cell><fo:block text-align="right">Fantastic Four</fo:block></fo:table-cell>
                    <fo:table-cell><fo:block> #240-253</fo:block></fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell><fo:block text-align="right">Fantastic Four Annual</fo:block></fo:table-cell>
                    <fo:table-cell><fo:block> #17</fo:block></fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell><fo:block text-align="right">The Avengers</fo:block></fo:table-cell>
                    <fo:table-cell><fo:block> #233</fo:block></fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
            </xsl:when>
            <xsl:when test="@volume_code='FFBYRNE_3'">

              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                 <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="center">John Byrne's Fantastic Four Volume 3</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
              
              <fo:block space-after=".25in" />


              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                  <fo:table-row>
                    <fo:table-cell><fo:block text-align="right">Fantastic Four</fo:block></fo:table-cell>
                    <fo:table-cell><fo:block> #254-267</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell><fo:block text-align="right">The Thing</fo:block>
                    </fo:table-cell>
                    <fo:table-cell>
                      <fo:block> #2, 10</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="right">Alpha Flight</fo:block>
                    </fo:table-cell>
                    <fo:table-cell>
                      <fo:block> #4</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
            </xsl:when>
            <xsl:when test="@volume_code='FFBYRNE_4'">

              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                 <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="center">John Byrne's Fantastic Four Volume 4</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
              
              <fo:block space-after=".25in" />



              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                  <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="right">Fantastic Four</fo:block>
                    </fo:table-cell>
                    <fo:table-cell>
                      <fo:block> #268-282</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="right">Fantastic Four Annual</fo:block>
                    </fo:table-cell>
                    <fo:table-cell>
                      <fo:block> #18</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell><fo:block text-align="right">The Thing</fo:block>
                    </fo:table-cell>
                    <fo:table-cell>
                      <fo:block> #19</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
            </xsl:when>
            <xsl:when test="@volume_code='FFBYRNE_5'">


              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                 <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="center">John Byrne's Fantastic Four Volume 5</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
              
              <fo:block space-after=".25in" />


              <fo:table table-layout="fixed" width="100%">
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-column column-width="proportional-column-width(1)"/>
                <fo:table-body>
                  <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="right">Fantastic Four</fo:block>
                    </fo:table-cell>
                    <fo:table-cell>
                      <fo:block> #283-295</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="right">Fantastic Four Annual</fo:block>
                    </fo:table-cell>
                    <fo:table-cell>
                      <fo:block> #19</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="right">The Thing</fo:block>
                    </fo:table-cell>
                    <fo:table-cell>
                      <fo:block> #23</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                  <fo:table-row>
                    <fo:table-cell>
                      <fo:block text-align="right">The Avengers Annual</fo:block>
                    </fo:table-cell>
                    <fo:table-cell>
                      <fo:block> #14</fo:block>
                    </fo:table-cell>
                  </fo:table-row>
                </fo:table-body>
              </fo:table>
            </xsl:when>
          </xsl:choose>

      </fo:block>


    <!--/td-->
      <!--/tr-->
    <!--/table-->

    <!--*** FOOTER ***-->
    <!--p class="footer" /-->
    <fo:block break-after='page'/>


  </xsl:template>

  <xsl:template name="issue_group">
    <xsl:param name="pos" />
    <!--table class="report" width="100%"-->
      <!--tr-->
        <!--td class="issue" width="50%"-->
    <fo:table table-layout="fixed" width="100%">
      <fo:table-column column-width="proportional-column-width(1)"/>
      <fo:table-column column-width="proportional-column-width(1)"/>
      <fo:table-body>
        <fo:table-row>
          <fo:table-cell>
            <fo:block>
            <!--POS:<xsl:value-of select="position()"/-->
            <xsl:call-template name="issue">
              <xsl:with-param name="pos">
                <xsl:value-of select="position()"/>
              </xsl:with-param>
            </xsl:call-template>
            <!--/td-->
            <!--td>XXXXX</td-->
              </fo:block>
          </fo:table-cell>
          <xsl:if test="position()+$MAX_ROWS &lt;= count(../issue)">
            <fo:table-cell>
              <fo:block>
              <!--td class="issue"-->
              <!--POS2:<xsl:value-of select="position()+$MAX_ROWS"/-->
              <xsl:call-template name="issue">
                <xsl:with-param name="pos">
                  <xsl:value-of select="position()+$MAX_ROWS"/>
                </xsl:with-param>
              </xsl:call-template>
                </fo:block>
                </fo:table-cell>
            <!--/td-->
          </xsl:if>
          <!--/tr-->
        </fo:table-row>
      </fo:table-body>
    </fo:table>
            <!--/table-->



  </xsl:template>

  <xsl:template name="issue">
    <xsl:param name="pos" />

    <xsl:variable name="issue_idx">
      <xsl:value-of select="$pos"/>
    </xsl:variable>

    <fo:block font-size="8pt" color="#FFFFFF" margin-left=".5in" margin-right=".5in" margin-top=".25in">

    <fo:block font-style="italic">
    <!--span class="issuetitle"-->
      "<xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Issue Title']" />"
    <!--/span--><!--br /-->
    </fo:block>

    <fo:block>
    <!--span class="issuename"-->
      <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Title']" /> #<xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Issue']" />
    <!--/span-->
    <fo:inline font-size="6pt"> (
    <xsl:call-template name="decodemonth">
      <xsl:with-param name="code">
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Mo']" />
      </xsl:with-param>
    </xsl:call-template> <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Year']" />)
    </fo:inline>
    </fo:block>
    <!--br /-->

    <xsl:if test="string-length(normalize-space(../issue[position()=$pos]/Fld[@Name='Writer'])) > 0">
      <fo:block>
      <fo:inline font-size="6pt">Writer: </fo:inline><!--span class="value"-->
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Writer']" />
      <!--/span--><!--br /-->
      </fo:block>
    </xsl:if>

    <xsl:if test="string-length(normalize-space(../issue[position()=$pos]/Fld[@Name='Artist'])) > 0">
      <fo:block>
      <fo:inline font-size="6pt">Artist: </fo:inline><!--span class="value"-->
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Artist']" />
      <!--/span--><!--br /-->
      </fo:block>
    </xsl:if>

    <xsl:if test="string-length(normalize-space(../issue[position()=$pos]/Fld[@Name='Penciller'])) > 0">
      <fo:block>
      <fo:inline font-size="6pt">Penciller: </fo:inline><!--span class="value"-->
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Penciller']" />
      <!--/span--><!--br /-->
      </fo:block>
    </xsl:if>

    <xsl:if test="string-length(normalize-space(../issue[position()=$pos]/Fld[@Name='Inker'])) > 0">
      <fo:block>
      <fo:inline font-size="6pt">Inker: </fo:inline><!--span class="value"-->
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Inker']" />
      <!--/span--><!--br /-->
      </fo:block>
    </xsl:if>

    <xsl:if test="string-length(normalize-space(../issue[position()=$pos]/Fld[@Name='Colorist'])) > 0">
      <fo:block>
      <fo:inline font-size="6pt">Colors: </fo:inline><!--span class="value"-->
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Colorist']" />
      <!--/span--><!--br /-->
      </fo:block>
    </xsl:if>

    <xsl:if test="string-length(normalize-space(../issue[position()=$pos]/Fld[@Name='Letterer'])) > 0">
      <fo:block>
      <fo:inline font-size="6pt">Letterer: </fo:inline><!--span class="value"-->
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Letterer']" />
      <!--/span--><!--br /-->
      </fo:block>
    </xsl:if>

    <xsl:if test="string-length(normalize-space(../issue[position()=$pos]/Fld[@Name='Cover Artist'])) > 0">
      <fo:block>
      <fo:inline font-size="6pt">Cover Artist: </fo:inline><!--span class="value"-->
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Cover Artist']" />
      <!--/span--><!--br /-->
      </fo:block>
    </xsl:if>

    <xsl:if test="string-length(normalize-space(../issue[position()=$pos]/Fld[@Name='Editor'])) > 0">
      <fo:block>
      <fo:inline font-size="6pt">Editor: </fo:inline><!--span class="value"-->
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Editor']" />
      <!--/span--><!--br /-->
      </fo:block>
    </xsl:if>

    <xsl:if test="string-length(normalize-space(../issue[position()=$pos]/Fld[@Name='Editor In Chief'])) > 0">
      <fo:block>
      <fo:inline font-size="6pt">Editor-In-Chief: </fo:inline><!--span class="value"-->
        <xsl:value-of select="../issue[position()=$pos]/Fld[@Name='Editor In Chief']" />
      <!--/span--><!--br /-->
      </fo:block>
    </xsl:if>

    <!--br /--><!--br /-->

    </fo:block>
    
  </xsl:template>

  <xsl:template name="decodemonth">
    <xsl:param name="code" />
    <xsl:choose>
      <xsl:when test="$code='1'">Jan.</xsl:when>
      <xsl:when test="$code='2'">Feb.</xsl:when>
      <xsl:when test="$code='3'">Mar.</xsl:when>
      <xsl:when test="$code='4'">Apr.</xsl:when>
      <xsl:when test="$code='5'">May</xsl:when>
      <xsl:when test="$code='6'">Jun.</xsl:when>
      <xsl:when test="$code='7'">Jul.</xsl:when>
      <xsl:when test="$code='8'">Aug.</xsl:when>
      <xsl:when test="$code='9'">Sep.</xsl:when>
      <xsl:when test="$code='10'">Oct.</xsl:when>
      <xsl:when test="$code='11'">Nov.</xsl:when>
      <xsl:when test="$code='12'">Dec.</xsl:when>
    </xsl:choose>
  </xsl:template>

</xsl:stylesheet>