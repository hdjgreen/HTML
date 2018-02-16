param(
# The folder that contains htm files, which is used for generating md files.
# For example: ..\DevDoc\Main\dev_data_binding\Build\markdown\dev_data_binding
[string]$htmFolder
)

# The flag indicate whether this htm file is changed, changed as $true.
[bool]$Script:modifyFlag = $false

# Represent the number of internal link in a htm file
[int]$Script:LinkNumber = 0

# Represent the number of heading associated with the <a id> tag in a htm file
[int]$Script:HeadingNumber = 0

# Represent the number of the cross link in a htm file
[int]$script:CrossLinkNumber = 0

# Entire link tag, eg:<a href="#in_this_topic">
[System.Collections.ArrayList]$Script:LinkFullcontents = @()

# The link ID relative to the above, eg: in_this_topic
[System.Collections.ArrayList]$Script:Linkcontents = @()

# Entire the content of the heading that contains <a id>
# eg: <h2><a id="In_this_topic"></a><a id="in_this_topic"></a><a id="IN_THIS_TOPIC"></a>In this topic</h2>
[System.Collections.ArrayList]$Script:Headingcontents = @()

# The <a id> tag, eg: <a id="In_this_topic"></a><a id="in_this_topic"></a><a id="IN_THIS_TOPIC"></a>
[System.Collections.ArrayList]$Script:IDcontents = @()

# Target Text of heading node associated with the <a id> tag, eg: In this topic
[System.Collections.ArrayList]$Script:Textcontents = @()

# Heading level eg: 1 or 2 or 3...
[System.Collections.ArrayList]$Script:HeadingLevel = @()

# Entire the cross link tag, eg: <a href="w8x_to_uwp_root.htm#if_you_have_an_8.1_universal_windows_app">
[System.Collections.ArrayList]$Script:CrossLink = @()

# The file name in the cross link, eg: w8x_to_uwp_root.htm
[System.Collections.ArrayList]$Script:CrossLinkFileName = @()

# The link ID in the cross link, eg: if_you_have_an_8.1_universal_windows_app
[System.Collections.ArrayList]$Script:CrossLinkContent = @()


# Get all htm files from this path
Function Get-HtmFile($path)
{
    # if HTM file location not as expected, then find HTM files under markdown folder and use its path
    if (-not (Test-Path $path))
    {
        $parentPath = $path.Substring(0, $path.LastIndexOf("\"))
         
        return @(Get-ChildItem -Path $parentPath -Filter *.htm -Recurse).FullName
    }

    return [system.IO.Directory]::GetFiles($path,'*.htm')
} 

# Edit all of the content that contain alert, internal link and table tags.
Function Edit-Content($content, $htm)
{

    $content = Fix-AlertTag $content

    $contentStr = $content | Out-String

    $contentStr = Fix-LinkAndHeadingTag $contentStr
    
    $contentStr = Fix-TableFormat $contentStr
    
    $contentStr = Fix-NobulletList $contentStr

    $contentStr = Fix-YAMLMetadata $contentStr $htm

    return $contentStr
}


Function Fix-AlertTag([System.Collections.ArrayList]$content = @())
{
    
    # Match div's start tag, there may be a string behind this, so adding .*
    $patternDivStart = [regex]'.*<div\s+class="alert">\s*<b>(.{3,15})</b>.*'

    # Match closest end tag
    $patternDivEnd = [regex]'.*</div>.*' 

    # Match the start tag and the end tag on the same line
    $patternDivEntire = [regex]'.*<div\s+class="alert">\s*<b>(.{3,15})</b>.*?</div>.*'

	# Initializaion as false, will be set to $true if matched alert div tag
    $start = $false  

	
    # store the conent in div tag
    [System.Collections.ArrayList]$contentDiv = @()
    # store the entire conent in the htm file
    [System.Collections.ArrayList]$contentNew = @()
        
    # Extract the content of the htm file, then check whether there are the div tag
    # of alert type, and it contains multi paragraphs inside.
	$content | ForEach-Object {
		# See whether the start tag and the end tag on the same line
		if($_ -match $patternDivEntire)
		{   

            $contentDiv.Add($_) | Out-Null

            [bool]$MultiPara = Edit-MultiPara $contentDiv
            Edit-DivTag $contentDiv $MultiPara

            $contentNew.AddRange($contentDiv)
            $contentDiv.Clear()

            $Script:modifyFlag = $true

		}
        # Matched div tag, and the contents of this div tag are divided into multiple lines
        elseif(($_ -notmatch $patternDivEntire) -and ($_ -match $patternDivStart)) 
        {
            $start = $true
            $contentDiv.Add($_) | Out-Null

            $Script:modifyFlag = $true

        }
        # Have started matching but not ending
        elseif(($_ -notmatch $patternDivEnd) -and $start) 
        {
            $contentDiv.Add($_) | Out-Null
        }
        # Have started matching and found end tag, then save this piece to array, and set $start = $false, terminate div
        elseif($start -and $_ -match $patternDivEnd)
        {
            $contentDiv.Add($_) | Out-Null

            [bool]$MultiPara = Edit-MultiPara $contentDiv
            Edit-DivTag $contentDiv $MultiPara
                
            $contentNew.AddRange($contentDiv)
            $contentDiv.Clear()
            $start = $false
        }
		else
		{
            $contentNew.Add($_) | Out-Null

		}

	}
    
    return $contentNew

}


# Input parameter is a array type, ruturn bool. $true indicates multi paragraph, otherwise $false.
# Replace it if is multi paragraph
Function Edit-MultiPara($Content)
{
    [bool]$IsMulti = $false
    $patternParaStart = [regex]'<p\s+class="(.{3,15})">'

    for($i = 0; $i -lt $Content.Count; $i++)
    {
        if($Content[$i] -match $patternParaStart)
        {
            # remove redundant new lines between notes with front \s*
            $old = "\s*<p\s+class=""{0}"">" -f $Matches[1]
            $new = "<p class=""{0}"">" -f $Matches[1]

            $Content[$i] = $Content[$i] -replace $old, $new

            $IsMulti = $true
        }       
    }

    return $IsMulti
}


# If is multi paragraph, $IsMulti as $true, not add GT sign after div
# If not, $IsMulti as $false, then add GT sign after div
Function Edit-DivTag($content, $IsMulti)
{
    $patternDivStart = [regex]'.*<div\s+class="alert">\s*<b>(.{3,15})</b>.*'
    
    $patternDivEnd = [regex]'</div>'

    $content[0] -match $patternDivStart | Out-Null

    # remove spaces of text with the last \s*
    $old = "<div\s+class=""alert"">\s*<b>{0}</b>\s*" -f $Matches[1]

    $newWithGT = "<div class=""alert""> <blockquote>[!{0}]<br/>" -f $Matches[1]

    $newWithOutGT = "<div class=""alert""> <blockquote>[!{0}]" -f $Matches[1]
    
    $newEnd = '</blockquote></div>'


    # The start tag must be the first in this array, so $content[0]
    if($IsMulti -eq $true) # multi paragraph
    {
        $content[0] = $content[0] -replace $old, $newWithOutGT 
    }
    else # not multi paragraphs
    {
        $content[0] = $content[0] -replace $old, $newWithGT 
    }

    # replace end div tag </div> with </blockquote></div>
    for($i = 0; $i -lt $Content.Count; $i++)
    {
        if($Content[$i] -match $patternDivEnd)
        {
            $content[$i] = $content[$i] -replace $patternDivEnd, $newEnd 
        }
    }
}


# Fix internal links
Function Fix-LinkAndHeadingTag($content)
{
   
    Get-LinkTag $content
    
    Get-HeadingTag $content
    
    $content = Set-Link $content

    $content = Set-Heading $content

    # Clean up relevant variables 
    $Script:LinkNumber = 0
    $Script:HeadingNumber = 0
    
    $Script:LinkFullcontents.Clear()
    $Script:Linkcontents.Clear()
    $Script:Headingcontents.Clear()
    $Script:IDcontents.Clear()
    $Script:Textcontents.Clear()
    $Script:HeadingLevel.Clear()

    return $content
}

# Get all internal link tags
Function Get-LinkTag($content)
{
    $patternLinkTag = [regex]'<a href="#(.*?)">'

    $content | Select-String $patternLinkTag -AllMatches | ForEach-Object {
    
        # The number of internal links found
        $Script:LinkNumber = $_.Matches.count
    
    
        foreach($v in $_.Matches)
        {
            # eg <a href="#in_this_topic">
            $value = $v.groups.value[0]
            $Script:LinkFullcontents.Add($value) | Out-Null           

            # eg in_this_topic
            $value = $v.groups.value[1]
            $Script:Linkcontents.Add($value) | Out-Null
    
        }
    }
}

# Get all heading tags that contain <a id>
Function Get-HeadingTag($content)
{
    $patternHeadingEntire = [regex]'<h([1-5])>(<a\s+id=.*?</a>)+([\s\S]*?)</h[1-5]>'

    $content | Select-String $patternHeadingEntire -AllMatches | ForEach-Object{
    
        # The number of heading with <a id> tags found
        $Script:HeadingNumber = $_.Matches.count
    
    
        foreach($v in $_.Matches)
        {
            # eg <h2><a id="In_this_topic"></a><a id="in_this_topic"></a><a id="IN_THIS_TOPIC"></a>In this topic</h2>
            $value = $v.groups.value[0]
            $Script:Headingcontents.Add($value) | Out-Null
    
            # Heading level eg. 1 or 2 or 3...
            $value = $v.groups.value[1]
            $Script:HeadingLevel.Add($value) | Out-Null
            
            # eg <a id="In_this_topic"></a><a id="in_this_topic"></a><a id="IN_THIS_TOPIC"></a>
            $value = $v.groups.value[2]
            $Script:IDcontents.Add($value) | Out-Null
            
            # eg In this topic
            $value = $v.groups.value[3]
            $Script:Textcontents.Add($value) | Out-Null
        }
    
    }
}

# Clean up all <a id> tags in headings whether or not it contians the corresponding link ID  
Function Set-Heading($content)
{
    for($j = 0; $j -lt $Script:HeadingNumber; $j++)
    {
        $oldHeading = $Script:Headingcontents[$j]
        
        # Processing the scenario of space-like symbols, this scenario will result in the internal link disable.
        $text = $Script:Textcontents[$j] -replace '\s', ' '
        $newHeading = '<h{0}>{1}</h{0}>' -f $Script:HeadingLevel[$j], $text
        
        $content = $content.Replace($oldHeading,$newHeading)

        $Script:modifyFlag = $true
    }

    return $content
}

# Find and replace with new link ID
Function Set-Link($content)
{
    for($i = 0; $i -lt $Script:LinkNumber; $i++)
    {
        for($j = 0; $j -lt $Script:HeadingNumber; $j++)
        {
            
            if($Script:Headingcontents[$j].Contains($Script:Linkcontents[$i]))
            {
                # extracted the heading without space at the beginning and end
                $trimString = $Script:Textcontents[$j].Trim()
                # First remove tags of target text of heading 
                # then convert spaces to hyphens and remove these chars ( ) / , @ \ ? ; : . ~ ! # $ % ^ & * + = ‘ ’ “ [ ] { } |
                # remain - _ and numbers
                $ID = ($trimString -replace '<[\w\W]*?>','' -replace '</[\w\W]*?>','' -replace "['’""\[\]\{\}\(\)@\\/,\?:\.;~!#\$%^&\*+=\|<>]",''`
                 -replace '\s{1,10}', '-').ToLower()
                
                # replace
                $oldLink = $Script:LinkFullcontents[$i]
                $newLink = '<a href="#{0}">' -f $ID

                $content = $content.Replace($oldLink,$newLink)

                $Script:modifyFlag = $true
    
            }
        }
    }

    return $content
}

Function  Fix-TableFormat($str)
{
    # Regex for table and divs
    $tabletag = [regex]'(?i)(<table.*?>[\s\S]*?</table>)'
    $divstartpattern = [regex]'<div\sclass="\w*">'

    $div = [regex]'<div>'
    $divend = [regex]'</div>'
    
    $p = [regex]'<p>'
    $pend = [regex]'</p>'

    # these need to be processed specially, no any content or only spaces 
    # between start and end tag
    $ptag = [regex]'(?i)(<p>\s*?</p>)'
    $divtag = [regex]'(?i)(<div>\s*?</div>)'

    $replaceString = '[NEWLINE]'

    #variable that holds array of all the matches of table tags
    [System.Collections.ArrayList]$tablecontents = @()
        
    # Extract the content of the htm file, then check whether there are div tags inside the table tag
    $strnew = $str # load the whole file as a single string

    $match = $tabletag.Match($str)
    while ($match.Success) {
        $tablecontents.Add($match.Value) | out-null
        $match = $match.NextMatch()
    } 
   
    foreach($tablecontent in $tablecontents)
    {

        $newtablecontent = $tablecontent
        # Remove these non-functional P and Div tags, such as <p></p> or <div>  </div>
        if(($tablecontent -match $ptag) -or ($tablecontent -match $divtag))
        {
            $newtablecontent = $newtablecontent -replace  $ptag, ''
            $newtablecontent = $newtablecontent -replace  $divtag, ''
        }

        $newtablecontent = $newtablecontent -replace  $divstartpattern, '' -replace $div, '' -replace $divend, $replaceString -replace $p, '' -replace $pend, $replaceString
        $strnew = $strnew.replace( $tablecontent,  $newtablecontent)
    }

    # update the file only if there are any changes
    if ($str -ne $strnew)
    {
	    $Script:modifyFlag = $true
    }

    return $strnew
}


Function Fix-CrossLinks($content, $htmPath)
{

    Get-CrossLinks $content

    for($i = 0; $i -lt $script:CrossLinkNumber; $i++)
    {
		$parentPath = $htmPath.Substring(0, $htmPath.LastIndexOf("\"))
	
        $targetFile = '{0}\{1}' -f $parentPath, $Script:CrossLinkFileName[$i]

        # Test the path. If the file doesn't exist under current path, ignore it.
        if(!(Test-Path $targetFile))
        {
            Write-Debug "The path $targetFile doesn't exist in $htmPath!"

            continue
        }

        $Str = Get-Content $targetFile -Encoding UTF8 | Out-String

        Get-HeadingTag $Str


        for($j = 0; $j -lt $Script:HeadingNumber; $j++)
        {
            if($Script:Headingcontents[$j].Contains($Script:CrossLinkContent[$i]))
            {
                # extracted the heading without space at the beginning and end
                $trimString = $Script:Textcontents[$j].Trim()
                # First remove tags of target text of heading 
                # then convert spaces to hyphens and remove these chars ( ) / , @ \ ? ; : . ~ ! # $ % ^ & * + = ‘ ’ “ [ ] { } |
                # remain - _ and numbers
                $ID = ($trimString -replace '<[\w\W]*?>','' -replace '</[\w\W]*?>','' -replace "['’""\[\]\{\}\(\)@\\/,\?:\.;~!#\$%^&\*+=\|<>]",''`
                 -replace '\s{1,10}', '-').ToLower()           

                $oldCrossLink = $Script:CrossLink[$i]
                $newCrossLink = '<a href="{0}#{1}">' -f $Script:CrossLinkFileName[$i], $ID

                $content = $content.Replace($oldCrossLink,$newCrossLink)
    
                $Script:modifyFlag = $true

            }
        }

        # Clean up the variables of the heading
        $Script:Headingcontents.Clear()
        $Script:IDcontents.Clear()
        $Script:Textcontents.Clear()
        $Script:HeadingLevel.Clear()
        $Script:HeadingNumber = 0

    }

    # Clean up the variables of the cross link
    $script:CrossLinkNumber = 0
    $Script:CrossLink.Clear()
    $Script:CrossLinkFileName.Clear()
    $Script:CrossLinkContent.Clear()

    return $content
}

# Get all of the cross link in a htm file
Function Get-CrossLinks($content)
{
    $patternCrossLinks = [regex]'<a\s+href="(\S*?\.htm)#(.*?)">'

    $content | Select-String $patternCrossLinks -AllMatches | ForEach-Object{
        
        $script:CrossLinkNumber = $_.Matches.count

        foreach($v in $_.Matches)
        {
             # Entire the cross link tag
             $value = $v.groups.value[0]
             $Script:CrossLink.Add($value) | Out-Null

             # The htm file name in the cross link
             $value = $v.groups.value[1]
             $Script:CrossLinkFileName.Add($value) | Out-Null

             # The link ID in the cross link
             $value = $v.groups.value[2]
             $Script:CrossLinkContent.Add($value) | Out-Null
        }
    }
}


Function Fix-NobulletList($content)
{
    # Match nested scenario
    $patternNobullet = [regex]"<dl[^>]*>[\s\S]*?(((?<Open><dl[^>]*>)[\s\S]*?)+((?<-Open></dl>)[\s\S]*?)+)*(?(Open)(?!))</dl>"
    $patterDTTag = [regex]'<dt>'
    $patternPTag = [regex]'<p>[\s\S]*?</p>'
    $patternDDTag = [regex]'(<dd>[\s\S]*?</dd>)'
    [string]$newString
    [int]$pNumber = 0

    $content | Select-String $patternNobullet -AllMatches | ForEach-Object{
        
        foreach($v in $_.Matches)
        {
            # Matched <dl> list items
            $dlContent = $v.groups.value[0]

            # Filter out list items that contain <dt> tag
            if($dlContent -notmatch $patterDTTag)
            {

                # Matched <dd> tags in each <dl>
                $dlContent | Select-String $patternDDTag -AllMatches | ForEach-Object{

                    foreach($v in $_.Matches)
                    {
                        $ddContent = $v.groups.value[0]

                        # check the numbere of <p> tag in a <dd> tag
                        # if greater than one <p> tag in a <dd> tag, multi-prargraph, need to replace </p> with <br/> and remove <dd>, </dd>, <p>
                        # if the number of <p> tag is less than or equal to 1, replace </dd> with <br/> directly, and remove <dd>, <p>, </p>
                        $ddContent | Select-String $patternPTag -AllMatches | ForEach-Object{ $pNumber = $_.Matches.count }

                        if($pNumber -gt 1)
                        {
                            $newString = $ddContent -replace '<dd>', '' -replace '</dd>', '' -replace '<p>', '' -replace '</p>', '<br/>'

                            $content = $content.Replace($ddContent,$newString)                    
                        }
                        else
                        {
                            $newString = $ddContent -replace '<dd>', '' -replace '</p>', '' -replace '<p>', '' -replace '</dd>', '<br/>'

                            $content = $content.Replace($ddContent,$newString)
                        }

                        $pNumber = 0

                        $Script:modifyFlag = $true
                    }
                }
            }
        }
    }

    return $content
}


Function Fix-YAMLMetadata($content, $htmPath)
{
    $patternDCS     = [regex]'<meta\s+name="DCS.appliesToProduct".*?\/>'
    $patternHAID    = [regex]'<meta\s+name="MS-HAID".*?\/>'
    $patternHAttr   = [regex]'<meta\s+name="MSHAttr".*?\/>'
    $patterndevlang = [regex]'<meta\s+name="ms\.devlang".*?\/>'
    $patternTopicID = [regex]'<meta\s+name="Search\.Refinement\.TopicID".*?\/>'
    $patternassetId = [regex]'<meta\s+name="ms\.assetid".*?\/>'
    $patternHead    = [regex]'<head>'

    $patternXMLassetId = [regex]'<metadata.*?msdnID\s*?=\s*?"(.*?)">'

    $PatternFileName = [regex]'.*\\(.*)\.htm'
    $patternExtractRootProject = [regex]'\\[Bb]uild\\.*'

    # remove metadate
    $newcontent = $content -replace $patternDCS, '' -replace $patternHAID, '' -replace $patternHAttr, '' -replace $patterndevlang, '' -replace $patternTopicID, ''

    # If the htm file desn't contain assetid, then we need to extract it from xml file
    if($content -notmatch $patternassetId)
    {
        # extract name without the suffix
        if($htmPath -match $PatternFileName)
        {
            $htmFileName = $Matches[1]

            # splicing to the project root path
            # similar to like this D:\EnlistmentPUB\DevDoc\Main\accessibility from D:\EnlistmentPUB\DevDoc\Main\accessibility\Build\markdown\accessibility\a_element.htm after replacing
            $patternRootProjectPath = $htmPath -replace $patternExtractRootProject, ''
            
            # traverse and find out all xml file in this project folder
            $XMLFileName = '{0}.xml' -f $htmFileName
            # May find two the same name xml files, their content is the same, so the first is ok.
            $XMLFilePath = @(@(Get-ChildItem -Path $patternRootProjectPath -Filter *.xml -Recurse).FullName | Where-Object {$_ -match $XMLFileName})[0]

            $contentXML = Get-Content $XMLFilePath -Encoding UTF8 | Out-String

            #<metadata id="fs.clfs_management_constants" type="ovw" msdnID="26b21653-d3e6-4f1f-9024-7b73fced9808">
            # extract the value of 'msdnID' in xml file if it has
            if($contentXML -match $patternXMLassetId)
            {
                $assetID = $Matches[1]

                # splicing to like <meta name="ms.assetid" content="40001833-2131-406A-AE86-49B9C55AAD38"/>
                $assetIDHtm = '<meta name="ms.assetid" content="{0}"/>' -f $assetID

                # insert into head tag in htm
                $newcontent = $newcontent -replace $patternHead, "$patternHead`n$assetIDHtm"
            }

        }
    }

    if($newcontent -ne $content)
    {
        $Script:modifyFlag = $true
    }
    
    return $newcontent
}

################################# Main ###############################
[string[]]$htmFiles = Get-HtmFile $htmFolder 

[int]$fileCount = 0
[int]$fileTotal = $htmFiles.Length

# Process cross link
foreach($htm in $htmFiles)
{
    
    $fileCount += 1
    $percent = ($fileCount / $fileTotal) * 100
    Write-Progress "Processing the cross links: $fileCount file of $fileTotal" -PercentComplete $percent

    $contentStr = Get-Content $htm -Encoding UTF8 | Out-String
    
    $contentStr = Fix-CrossLinks $contentStr $htm

    if($Script:modifyFlag -eq $true)
    {
        Set-Content -Path $htm -Value $contentStr -Encoding UTF8 -NoNewline -Force

        $Script:modifyFlag = $false

        Write-Debug ('This htm file has been modified because of the cross links: {0}' -f $htm)
    }
}
 
$fileCount = 0

foreach($htm in $htmFiles)
{

    $fileCount += 1
    $percent = ($fileCount / $fileTotal) * 100
    Write-Progress "Processing other issues: $fileCount file of $fileTotal" -PercentComplete $percent

    $contentFull = Get-Content $htm -Encoding UTF8

    $contentFullStr = Edit-Content $contentFull $htm

    if($Script:modifyFlag -eq $true)
    {
        Set-Content -Path $htm -Value $contentFullStr -Encoding UTF8 -NoNewline -Force

        $Script:modifyFlag = $false

        Write-Debug ('This htm file has been modified: {0}' -f $htm)
    }
}
######################################################################