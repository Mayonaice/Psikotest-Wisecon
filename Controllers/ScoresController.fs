namespace PsikotestWisesa.Controllers

open System
open System.Data
open System.Text
open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authorization
open Microsoft.Data.SqlClient

[<CLIMutable>]
type ModifyReq = { Kind: string; Id: Nullable<int64>; Question: string; Aspek: string; Uraian: string; User: string }

type ScoresController (db: IDbConnection) =
    inherit Controller()

    member private this.mapTable(kind: string) =
        let k = if String.IsNullOrEmpty(kind) then "" else kind
        match k.Trim().ToLowerInvariant() with
        | "personal" | "datapribadi" -> "WISECON_PSIKOTEST.dbo.REC_MsPersonalInterview"
        | "experience" | "pengalaman" -> "WISECON_PSIKOTEST.dbo.REC_MsExperienceInterview"
        | "personality" | "kepribadian" -> "WISECON_PSIKOTEST.dbo.REC_MsPersonalityTest"
        | _ -> null

    member private this.getColumns(conn: SqlConnection, fullTableName: string) =
        use cmd = new SqlCommand("SELECT COLUMN_NAME, DATA_TYPE FROM " + fullTableName.Substring(0, fullTableName.IndexOf('.')) + ".INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='dbo' AND TABLE_NAME=@t")
        cmd.Parameters.AddWithValue("@t", fullTableName.Substring(fullTableName.LastIndexOf('.')+1)) |> ignore
        cmd.Connection <- conn
        use rdr = cmd.ExecuteReader()
        let cols = ResizeArray<string * string>()
        while rdr.Read() do
            cols.Add(rdr.GetString(0), rdr.GetString(1))
        cols |> Seq.toList

    member private this.detectIdColumn(columns: (string*string) list) =
        let candidates = ["SeqNo"; "ID"; "Id"; "NoSeq"; "NoUrut"; "No" ]
        candidates |> List.tryFind (fun c -> columns |> List.exists (fun (nm,_) -> String.Equals(nm, c, StringComparison.OrdinalIgnoreCase)))
        |> Option.defaultValue (fst (columns |> List.tryFind (fun (_,dt) -> dt.StartsWith("int")) |> Option.defaultValue ("SeqNo","int")))

    member private this.detectQuestionColumn(columns: (string*string) list) =
        let candidates = ["Question"; "Pertanyaan"; "Deskripsi"; "Keterangan"; "Nama"; "Judul"; "Notes" ]
        candidates |> List.tryFind (fun c -> columns |> List.exists (fun (nm,_) -> String.Equals(nm, c, StringComparison.OrdinalIgnoreCase)))
        |> Option.defaultValue (fst (columns |> List.tryFind (fun (_,dt) -> dt.Contains("char") || dt.Contains("text")) |> Option.defaultValue ("Question","nvarchar")))

    member private this.hasColumn(columns: (string*string) list, name: string) =
        columns |> List.exists (fun (nm,_) -> String.Equals(nm, name, StringComparison.OrdinalIgnoreCase))

    member private this.isPersonalityKind(kind: string) =
        let k = if String.IsNullOrEmpty(kind) then "" else kind
        match k.Trim().ToLowerInvariant() with
        | "personality" | "kepribadian" -> true
        | _ -> false

    [<Authorize>]
    [<HttpGet>]
    [<Route("Scores/List")>]
    member this.List([<FromQuery>] kind: string, [<FromQuery>] start: Nullable<DateTime>, [<FromQuery>] endd: Nullable<DateTime>) : IActionResult =
        let table = this.mapTable(kind)
        if String.IsNullOrWhiteSpace(table) then this.BadRequest(box {| error = "Unknown kind" |}) :> IActionResult
        else
            let conn = db :?> SqlConnection
            conn.Open()
            try
                let cols = this.getColumns(conn, table)
                let idCol = this.detectIdColumn(cols)
                let qCol = this.detectQuestionColumn(cols)
                let hasUser = this.hasColumn(cols, "UserInput")
                let hasTime = this.hasColumn(cols, "TimeInput")
                let isPersonality = this.isPersonalityKind(kind)
                let hasAspek = isPersonality && this.hasColumn(cols, "Aspek")
                let hasUraian = isPersonality && this.hasColumn(cols, "Uraian")
                let mutable whereClause = " WHERE 1=1 "
                use cmd = new SqlCommand()
                cmd.Connection <- conn
                cmd.CommandType <- CommandType.Text
                if hasTime && start.HasValue && endd.HasValue then
                    whereClause <- whereClause + " AND CAST(TimeInput AS DATE) BETWEEN @s AND @e "
                    cmd.Parameters.AddWithValue("@s", start.Value.Date) |> ignore
                    cmd.Parameters.AddWithValue("@e", endd.Value.Date) |> ignore
                let selectUser = if hasUser then "UserInput" else "'' AS UserInput"
                let selectTime = if hasTime then "TimeInput" else "GETDATE() AS TimeInput"
                let selectQuestion =
                    if hasUraian then "ISNULL([Uraian],'') AS Question"
                    else sprintf "ISNULL(%s,'') AS Question" qCol
                let selectAspek = if hasAspek then "ISNULL([Aspek],'') AS Aspek" else "'' AS Aspek"
                let selectUraian = if hasUraian then "ISNULL([Uraian],'') AS Uraian" else "'' AS Uraian"
                cmd.CommandText <- sprintf "SELECT %s AS Id, %s, %s, %s, %s, %s FROM %s%s ORDER BY %s DESC" idCol selectQuestion selectAspek selectUraian selectUser selectTime table whereClause idCol
                use rdr = cmd.ExecuteReader()
                let out = ResizeArray<obj>()
                let mutable i = 0
                while rdr.Read() do
                    i <- i + 1
                    let id = rdr.GetValue(0)
                    let q = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                    let aspek = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                    let uraian = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                    let user = if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                    let time = if rdr.IsDBNull(5) then Nullable() else Nullable(rdr.GetDateTime(5))
                    out.Add(box {| num = i; id = id; question = q; aspek = aspek; uraian = uraian; userInput = user; timeInput = time |})
                this.Ok(out)
            finally
                conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Scores/Add")>]
    member this.Add([<FromBody>] body: ModifyReq) : IActionResult =
        let table = this.mapTable(body.Kind)
        if String.IsNullOrWhiteSpace(table) then this.BadRequest(box {| error = "Unknown kind" |}) :> IActionResult
        else
            let conn = db :?> SqlConnection
            conn.Open()
            try
                let cols = this.getColumns(conn, table)
                let qCol = this.detectQuestionColumn(cols)
                let hasUser = this.hasColumn(cols, "UserInput")
                let hasTime = this.hasColumn(cols, "TimeInput")
                let isPersonality = this.isPersonalityKind(body.Kind)
                let hasAspek = isPersonality && this.hasColumn(cols, "Aspek")
                let hasUraian = isPersonality && this.hasColumn(cols, "Uraian")
                let columnsList = ResizeArray<string>()
                let valuesList = ResizeArray<string>()
                use cmd = new SqlCommand()
                cmd.Connection <- conn
                cmd.CommandType <- CommandType.Text
                if hasAspek && hasUraian then
                    columnsList.Add("Aspek")
                    valuesList.Add("@a")
                    let av = if isNull body.Aspek then "" else body.Aspek
                    cmd.Parameters.AddWithValue("@a", av) |> ignore

                    columnsList.Add("Uraian")
                    valuesList.Add("@r")
                    let qv = if obj.ReferenceEquals(body.Question, null) then "" else body.Question
                    let rv = if isNull body.Uraian then qv else body.Uraian
                    cmd.Parameters.AddWithValue("@r", rv) |> ignore
                else
                    columnsList.Add(qCol)
                    valuesList.Add("@q")
                    let qv = if obj.ReferenceEquals(body.Question, null) then "" else body.Question
                    cmd.Parameters.AddWithValue("@q", qv) |> ignore
                if hasUser then
                    columnsList.Add("UserInput")
                    valuesList.Add("@u")
                    let uv = if obj.ReferenceEquals(body.User, null) then "" else body.User
                    cmd.Parameters.AddWithValue("@u", uv) |> ignore
                if hasTime then columnsList.Add("TimeInput"); valuesList.Add("GETDATE()")
                cmd.CommandText <- sprintf "INSERT INTO %s (%s) VALUES (%s)" table (String.Join(",", columnsList)) (String.Join(",", valuesList))
                let _ = cmd.ExecuteNonQuery()
                this.Ok(box {| ok = true |})
            finally
                conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Scores/Edit")>]
    member this.Edit([<FromBody>] body: ModifyReq) : IActionResult =
        let table = this.mapTable(body.Kind)
        if String.IsNullOrWhiteSpace(table) then this.BadRequest(box {| error = "Unknown kind" |}) :> IActionResult
        elif not body.Id.HasValue then this.BadRequest(box {| error = "Missing Id" |}) :> IActionResult
        else
            let conn = db :?> SqlConnection
            conn.Open()
            try
                let cols = this.getColumns(conn, table)
                let idCol = this.detectIdColumn(cols)
                let qCol = this.detectQuestionColumn(cols)
                let hasUE = this.hasColumn(cols, "UserEdit")
                let hasTE = this.hasColumn(cols, "TimeEdit")
                let isPersonality = this.isPersonalityKind(body.Kind)
                let hasAspek = isPersonality && this.hasColumn(cols, "Aspek")
                let hasUraian = isPersonality && this.hasColumn(cols, "Uraian")
                use cmd = new SqlCommand()
                cmd.Connection <- conn
                cmd.CommandType <- CommandType.Text
                let mutable setClause = ""
                if hasAspek && hasUraian then
                    setClause <- "Aspek=@a, Uraian=@r"
                    let av = if isNull body.Aspek then "" else body.Aspek
                    cmd.Parameters.AddWithValue("@a", av) |> ignore
                    let qv = if obj.ReferenceEquals(body.Question, null) then "" else body.Question
                    let rv = if isNull body.Uraian then qv else body.Uraian
                    cmd.Parameters.AddWithValue("@r", rv) |> ignore
                else
                    setClause <- sprintf "%s=@q" qCol
                    let qv = if obj.ReferenceEquals(body.Question, null) then "" else body.Question
                    cmd.Parameters.AddWithValue("@q", qv) |> ignore
                if hasUE then
                    setClause <- setClause + ", UserEdit=@ue"
                    let uv = if obj.ReferenceEquals(body.User, null) then "" else body.User
                    cmd.Parameters.AddWithValue("@ue", uv) |> ignore
                if hasTE then setClause <- setClause + ", TimeEdit=GETDATE()"
                cmd.CommandText <- sprintf "UPDATE %s SET %s WHERE %s=@id" table setClause idCol
                cmd.Parameters.AddWithValue("@id", body.Id.Value) |> ignore
                let _ = cmd.ExecuteNonQuery()
                this.Ok(box {| ok = true |})
            finally
                conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Scores/Delete")>]
    member this.Delete([<FromBody>] body: ModifyReq) : IActionResult =
        let table = this.mapTable(body.Kind)
        if String.IsNullOrWhiteSpace(table) then this.BadRequest(box {| error = "Unknown kind" |}) :> IActionResult
        elif not body.Id.HasValue then this.BadRequest(box {| error = "Missing Id" |}) :> IActionResult
        else
            let conn = db :?> SqlConnection
            conn.Open()
            try
                let cols = this.getColumns(conn, table)
                let idCol = this.detectIdColumn(cols)
                use cmd = new SqlCommand(sprintf "DELETE FROM %s WHERE %s=@id" table idCol, conn)
                cmd.Parameters.AddWithValue("@id", body.Id.Value) |> ignore
                let _ = cmd.ExecuteNonQuery()
                this.Ok(box {| ok = true |})
            finally
                conn.Close()
