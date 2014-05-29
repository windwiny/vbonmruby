Attribute VB_Name = "Module_vbonmruby"
Option Explicit

'change this value if mruby code need retun more len char[]
Public Const MAX_RETURN_LEN = 4096

'random uuid, used by join params string
Public Const SPLIT_STR = "72e525c3-6958-44b7-8293-2b133d88df84"



'return a long. 1: 1 result, 2: string joined by SPLIT_STR, -1: UNDEF syntax error. 0: other
'ruby run result value in result_str, using parsm_spit_str split it.
'On ruby code, using $paramss split $params

Public Declare Function vbonmruby_load_string Lib "vbonmruby.dll" ( _
        ByVal run_str As String, _
        ByVal param_str As String, _
        ByVal param_split_str As String, _
        ByVal result_str As String, _
        ByVal uSize As Long _
    ) As Long


Public Function VB_run(runstr As String, ParamArray Params() As Variant) As String()
  Dim pars As String
  Dim x As Variant
  
  'Escape and Join Parameters string
  pars = ""
  For Each x In Params
    x = CStr(x)
    'x = Replace(x, """", "\""")             'replace " to \"
    'x = Replace(x, "'", "\'")               'replace ' to \'
    'pars = pars & "  " & """" & x & """"    'every parameter included in two double quote(") char
    If pars = "" Then
      pars = x
    Else
      pars = pars & SPLIT_STR & x
    End If
  Next
  
  'MsgBox runstr, vbOKOnly, "runstr"
  'MsgBox pars, vbOKOnly, "pars"
  
  Dim Result As String
  Dim result_len As Long
  Result = String(MAX_RETURN_LEN, Chr(0))

  result_len = vbonmruby_load_string(runstr, pars, SPLIT_STR, Result, MAX_RETURN_LEN)

  Dim sss() As String
  'If result_len <> -1 Then
  sss = Split(Result, SPLIT_STR)
  'End If
  VB_run = sss
End Function

