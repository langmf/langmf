<#Module=libFile>
'------------------------------------------------
Sub GO()
  info_my message!@#$%^&*()_+
End Sub
<#Module>

<#Script=Parser2,VBScript#>
  function LMF_Parser_Raw(code) : Parser2.Sys.Ext.msgbox(mf_IDS) : LMF_Parser_Raw = code : end function
<#Script#>

<#Script=Parser3#>
  function LMF_Parser_Raw(code) {	Parser3.Sys.Ext.msgbox(mf_IDS); return code.replace(/info_([^\r]*)/, 'call msgbox("$1",vbcritical,"info")')}
<#Script#>