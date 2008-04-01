<?xml version="1.0" ?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">

	<xsl:output method="xml" omit-xml-declaration="yes" indent="yes"/>
	<xsl:template match="*">

		<xsl:text disable-output-escaping="yes"></xsl:text>
		
		<xsl:value-of select="*[local-name()='channel']/*[local-name()='lastBuildDate']"/>
		
		<div>
		
		<xsl:for-each select="//*[local-name()='item']">
		<xsl:if test="position() &lt; 6">
		
			<div style="padding:4px;" class="small">
		
			<a>
			
			<xsl:attribute name="href">
				<xsl:value-of select="*[local-name()='link']"/>
			</xsl:attribute>
			
			<xsl:attribute name="target">
				<xsl:text>top</xsl:text>
			</xsl:attribute>
			
			<xsl:value-of select="*[local-name()='title']"/>
			
			</a>
			
			<xsl:value-of select="*[local-name()='description']" disable-output-escaping="yes"/>
		
			</div>
		
		</xsl:if>
		
		</xsl:for-each>
		
		</div>
		
	</xsl:template>
	
	<xsl:template match="/">
	
		<xsl:apply-templates/>
		
	</xsl:template>
	
</xsl:stylesheet>
