Declare Function sqlite3_create_function Lib "*sqlite3" (byval db as long, byval zName as string, byval nArg as long, byval eTextRep as long, byval pApp as long, byval xFunc as long, byval xStep as long, byval xFinal as long) as long
Declare Function sqlite3_exec Lib "*sqlite3" (byval db as long, byval sql as string, byval callback as long, byval arg1 as long, err as long) As long
Declare Function sqlite3_get_table Lib "*sqlite3" (byval db as long, byval sql as string, result as long, row as long, col as long, err as long) As long
Declare Function sqlite3_open Lib "*sqlite3" (byval name as string, db as long) As long
Declare Function sqlite3_close Lib "*sqlite3" (byval db as long) As long
Declare Function sqlite3_free Lib "*sqlite3" (byval ptr as long) As long
Declare Function sqlite3_free_table Lib "*sqlite3" (byval ptr as long) As long
Declare Function sqlite3_errmsg Lib "*sqlite3" (byval db as long) As String
Declare Function sqlite3_errcode Lib "*sqlite3" (byval db as long) As long
Declare Function sqlite3_value_text Lib "*sqlite3" (byval value as long) as string
Declare Function sqlite3_libversion Lib "*sqlite3" () As String
Declare Sub sqlite3_result_int Lib "*sqlite3" (byval context as long, byval int as long)

<#Module=mdlSQLite>
Class SQLite

    Dim hDB, Fields(), CBK(), RXP

    '-------------------------------------
    Private Sub Class_Initialize()
        Set RXP = New RegExp   :   RXP.IgnoreCase = True
    End Sub

    Function CreateFunction(realFunc, nameFunc, numArg)
        Dim i   :   i = ArraySize(CBK)   :   Redim Preserve CBK(i)
        Set CBK(i) = sys.callback(realFunc, Me, 3, True)    :   sqlite3_create_function hDB, nameFunc, numArg, 0, 0, CBK(i), 0, 0
    End Function

    '-------------------------------------    
    Function cbk_REGEXP(ctx, cnt, pArg)
        RXP.Pattern = sqlite3_value_text(Struct(pArg + 0))  :   sqlite3_result_int ctx, RXP.Test(sqlite3_value_text(Struct(pArg + 4)))
    End Function

    '-------------------------------------
    Function Close()    :   Close = sqlite3_close(hDB)                  :   End Function
    Function Exec(vSQL) :   Exec = sqlite3_exec(hDB, vSQL, 0, 0, 0)     :   End Function
    Function ErrCode()  :   ErrCode = sqlite3_errcode(hDB)              :   End Function
    Function ErrMsg()   :   ErrMsg = sqlite3_errmsg(hDB)                :   End Function

    Function Open(nameDB)
        Open = sqlite3_open(nameDB, hDB)
        If Open = 0 Then
            Exec "PRAGMA journal_mode = memory"
            CreateFunction "cbk_REGEXP", "REGEXP", 2
        End If
    End Function

    Function Query(vSQL)
        Dim c, r, ptr, rows, cols, tmp, dat

        Erase Fields   :   Set Query = sys.newcollection

        If sqlite3_get_table(hDB, vSQL, ptr, rows, cols, 0) = 0 Then
                rows = rows + 1     :    Redim Fields(cols - 1)     :    dat = sys.conv.ptr2str(ptr,,cols,rows)
                
                For c = 0 To cols - 1	:	Fields(c) = dat(c,0)    :   Next
                
                For r = 1 To rows - 1
                    Set tmp = sys.newcollection
                    With tmp   :   For c = 0 To cols - 1    :   .add dat(c,r), Fields(c)    :   Next    :   End With
                    Query.Add tmp
                Next
        End If

        sqlite3_free_table ptr
    End Function

End Class
<#Module>