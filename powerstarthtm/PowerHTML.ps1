class PowerStartHTML {
    [string]$PowerStartHtmlTemplate = "<html><head><meta charset=`"UTF-8`"/><title></title><style></style></head><body><div/></body></html>"
    [xml]$xmlDocument = $null
    $cssStyles= @{}
    $lastEl = $null
    $newEl = $null
    PowerStartHTML($title) {
        $this.xmlDocument = $this.PowerStartHtmlTemplate
        $this.xmlDocument.html.head.title = $title
        $this.lastEl = $this.xmlDocument.html.body.ChildNodes[0]
    }
    [string] GetHtml() {
        $csb = [System.Text.StringBuilder]::new()
        foreach ($cssStyle in $this.cssStyles.GetEnumerator()) {
            $null = $csb.AppendFormat("{0} {{ {1} }}",$cssStyle.Name,$cssStyle.Value)
        }
        $this.xmlDocument.html.head.style = $csb.toString()
        return  ("<!DOCTYPE html>{0}" -f $this.xmlDocument.OuterXML)
    }
    Save($path) {
        $this.GetHtml() | set-content -path $path
    }
    AddAttr($el,$name,$value) {
        $attr = $this.xmlDocument.CreateAttribute($name)
        $attr.Value = $value
        $el.Attributes.Append($attr)
    }
    AddAttrs($el,$dict) {
        foreach($a in $dict.GetEnumerator()) {
            $this.AddAttr($el,$a.Name,$a.Value)
        }
    }
     [PowerStartHTML] AddBootStrap() {
        # <link rel="stylesheet" href="" integrity="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u" >
        $el = $this.xmlDocument.CreateElement("link")
        $attrs = @{
            "rel"="stylesheet";
            "href"="https://maxcdn.bootstrapcdn.com/bootstrap/3.3.7/css/bootstrap.min.css";
            "integrity"="sha384-BVYiiSIFeK1dGmJRAkycuHAHRg32OmUcww7on3RYdg4Va+PmSTsz/K68vbdEjh4u";
            "crossorigin"="anonymous";
        }
        $this.AddAttrs($el,$attrs)
        $this.xmlDocument.html.head.AppendChild($el)
        return $this
    }
     [PowerStartHTML]  AddContainerAttrToMain() {
        $this.AddAttr($this.xmlDocument.html.body.ChildNodes[0],"class","container")
        return $this
    }
    [PowerStartHTML] Append($elType = "table",$className=$null,[string]$text=$null) {
        $el = $this.xmlDocument.CreateElement($elType)
        if($text -ne $null) {
            $el.AppendChild($this.xmlDocument.CreateTextNode($text))
        } 
        if($className -ne $null) {
            $this.AddAttr($el,"class",$className)
        }
        $this.lastEl.AppendChild($el)
        $this.newEl = $el

        return $this
    }
    [PowerStartHTML] Append($elType = "table",$className=$null) { return $this.Append($elType,$className,$null) }
    [PowerStartHTML] Append($elType = "table") { return $this.Append($elType,$null,$null) }
    [PowerStartHTML] Add($elType = "table",$className=$null,[string]$text=$null) {
        $this.Append($elType,$className,$text)
        $this.lastEl = $this.newEl
        return $this
    }
    [PowerStartHTML] Add($elType = "table",$className=$null) { return $this.Add($elType,$className,$null) }
    [PowerStartHTML] Add($elType = "table") { return $this.Add($elType,$null,$null) }
    [PowerStartHTML] Main() {
        $this.lastEl = $this.xmlDocument.html.body.ChildNodes[0];
        return $this
    }
    [PowerStartHTML] Up() {
        $this.lastEl = $this.lastEl.ParentNode;
        return $this
    }
    N() {}
}
class PowerStartHTMLPassThroughLine {
    $object;$cells
    PowerStartHTMLPassThroughLine($object) {
        $this.object = $object; 
        $this.cells = new-object System.Collections.HashTable;
    }
}
class PowerStartHTMLPassThroughElement {
    $name;$text;$element
    PowerStartHTMLPassThroughElement($name,$text,$element) {
        $this.name = $name; $this.text = $text; $this.element = $element
    }
}

function Add-PowerStartHTMLTable {
    param(
         [Parameter(Mandatory=$True,ValueFromPipeline=$True)]$object,
         [PowerStartHTML]$psHtml,
         [string]$tableTitle = $null,
         [string]$tableClass = $null,
         [string]$idSuffix = $null,
         [switch]$passthroughTable = $false
    )
    begin {
        if($tableTitle -ne $null) {
            $psHtml.Main().Append("h1",$null,$tableTitle).N()
            if($idSuffix -ne $null) {
                $psHtml.AddAttr($psHtml.newEl,"id","header-$idsuffix")
            }   
        } 
        $psHtml.Main().Add("table").N()
        if($idSuffix -ne $null) {
           $psHtml.AddAttr($psHtml.newEl,"id","table-$idsuffix")
        }      
        if($tableClass -ne $null) {
           $psHtml.AddAttr($psHtml.newEl,"class",$tableClass)
        }     
    }
    process {
        $psHtml.Add("tr").N()
        $pstableln = [PowerStartHTMLPassThroughLine]::new($object)
        $object | Get-Member -Type Properties | % {
            $n = $_.Name;
            $psHtml.Append("td",$null,$object."$n").N()
            if($passthroughTable) {
                $pstableln.cells.Add($n,[PowerStartHTMLPassThroughElement]::new($n,($object."$n"),$psHtml.newEl))
            }
        }
        if($passthroughTable) {
            $pstableln
        }
        $psHtml.Up().N()
    }
    end { 
    }
}







