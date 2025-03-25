Add-Type -AssemblyName System.Windows.Forms

function Open-Webpage($url){

    [System.Windows.Forms.Application]::EnableVisualStyles()

    $Form                            = New-Object System.Windows.Forms.Form
    $Form.ClientSize                 = New-Object System.Drawing.Point(621,520)
    $Form.text                       = $title
    $Form.TopMost                    = $false
    $Form.Icon                       = [System.Drawing.SystemIcons]::Exclamation
    $WebBrowser                      = New-Object system.Windows.Forms.WebBrowser
    $WebBrowser.width                = $form.Width
    $WebBrowser.height               = $form.height
    $WebBrowser.location             = New-Object System.Drawing.Point(2,4)
    $Form.controls.AddRange(@($WebBrowser))

    $Form.Add_Load({  
        $WebBrowser.Url = $url
    })

    $Form.Add_SizeChanged({ 
        $WebBrowser.width                = $form.Width
        $WebBrowser.height               = $form.height
    })
    [System.Windows.Forms.Application]::Run($form)

    $form.dispose()
}

