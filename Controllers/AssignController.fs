namespace PsikotestWisesa.Controllers

open System
open System.Data
open System.IO
open System.Text
open System.Security.Cryptography
open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authorization
open Microsoft.Extensions.Configuration
open ClosedXML.Excel

[<CLIMutable>]
type UjianFilterRequest = { inputStart: Nullable<DateTime>; inputEnd: Nullable<DateTime>; ujianStart: Nullable<DateTime>; ujianEnd: Nullable<DateTime>; status: string; hasil: string }

[<CLIMutable>]
type AssignUjianRequest = { Batch: string; NoPeserta: int64; WaktuUjian: DateTime; NoPaket: int64; User: string }

[<CLIMutable>]
type AssignUjianBulkRequest = { Batch: string; NoPeserta: int64 array; WaktuUjian: DateTime; NoPaket: int64; User: string }

[<CLIMutable>]
type EditUjianRequest = { NoPaket: int64; UserID: string; WaktuUjian: DateTime; block: bool; User: string }

[<CLIMutable>]
type ResendUjianRequest = { UserID: string; User: string }

[<CLIMutable>]
type InterviewFilterRequest = { inputStart: Nullable<DateTime>; inputEnd: Nullable<DateTime>; statusInterview: string }

[<CLIMutable>]
type AssignInterviewRequest = { Batch: string; NoPeserta: int64; WaktuInterview: DateTime; Lokasi: string; User: string }

[<CLIMutable>]
type AssignInterviewBulkRequest = { Batch: string; NoPeserta: int64 array; WaktuInterview: DateTime; Lokasi: string; User: string }

[<CLIMutable>]
type EditInterviewRequest = { SeqNo: int64; Batch: string; WaktuInterview: DateTime; Lokasi: string; User: string }

[<CLIMutable>]
type ResendInterviewRequest = { SeqNo: int64; User: string }

[<ApiController>]
type AssignController (db: IDbConnection, cfg: IConfiguration) =
    inherit Controller()

    let encrypt64 (plainText: string) (encryptionKey: string) =
        if String.IsNullOrWhiteSpace(plainText) || String.IsNullOrWhiteSpace(encryptionKey) then
            ""
        else
            let key8 =
                if encryptionKey.Length >= 8 then encryptionKey.Substring(0, 8) else ""

            if String.IsNullOrWhiteSpace(key8) then
                ""
            else
                use des = new DESCryptoServiceProvider()
                let keyBytes = Encoding.UTF8.GetBytes(key8)
                let ivBytes = [| 0x12uy; 0x34uy; 0x56uy; 0x78uy; 0x90uy; 0xABuy; 0xCDuy; 0xEFuy |]
                let input = Encoding.UTF8.GetBytes(plainText)
                use encryptor = des.CreateEncryptor(keyBytes, ivBytes)
                let cipher = encryptor.TransformFinalBlock(input, 0, input.Length)
                Convert.ToBase64String(cipher)

    let getSetting (keys: string list) =
        keys
        |> List.tryPick (fun k ->
            let v = cfg.[k]
            if String.IsNullOrWhiteSpace(v) then None else Some v)
        |> Option.defaultValue ""

    let getDomainName () =
        getSetting [ "Psikotest:DomainName"; "DomainName" ]

    let getEncryptionKey () =
        getSetting [ "Psikotest:EncryptionKey"; "EncryptionKey" ]

    let getPathBase () =
        getSetting [ "Psikotest:PathBase"; "PathBase" ]

    let buildExamUrl (domainName: string) (pathBase: string) (tokenEnc: string) =
        let tokenParam = System.Net.WebUtility.UrlEncode(tokenEnc)
        let baseDomain = if String.IsNullOrWhiteSpace(domainName) then "" else domainName.Trim()
        let pb =
            if String.IsNullOrWhiteSpace(pathBase) then ""
            else
                let t = pathBase.Trim()
                if t.StartsWith("/", StringComparison.Ordinal) then t else "/" + t

        let mutable finalBase = baseDomain
        let mutable finalPathBase = pb

        let mutable uri = Unchecked.defaultof<Uri>
        if Uri.TryCreate(baseDomain, UriKind.Absolute, &uri) then
            if not (String.IsNullOrWhiteSpace(pb)) then
                let pbNorm = pb.TrimEnd('/').ToLowerInvariant()
                let pathNorm = uri.AbsolutePath.TrimEnd('/').ToLowerInvariant()
                if pathNorm.EndsWith(pbNorm, StringComparison.OrdinalIgnoreCase) then
                    finalPathBase <- ""

        let baseUrl =
            if finalBase.EndsWith("/", StringComparison.Ordinal) then finalBase.TrimEnd('/')
            else finalBase

        let urlPathBase =
            if String.IsNullOrWhiteSpace(finalPathBase) then ""
            else finalPathBase.TrimEnd('/')

        baseUrl + urlPathBase + "/TRX_Ujian?Token=" + tokenParam

    [<Authorize>]
    [<HttpGet>]
    [<Route("Applicants/PaketSoal/List")>]
    member this.ListPaketSoal() : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT NoPaket, NamaPaket FROM WISECON_PSIKOTEST.dbo.MS_PaketSoal WHERE ISNULL(bAktif,0)=1 ORDER BY NoPaket"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let noPaket = if rdr.IsDBNull(0) then 0L else Convert.ToInt64(rdr.GetValue(0))
                let namaPaket = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                rows.Add(box {| noPaket = noPaket; namaPaket = namaPaket |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Applicants/Batch/List")>]
    member this.ListBatch() : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <-
            "SELECT Batch FROM (" +
            "  SELECT DISTINCT Batch FROM WISECON_PSIKOTEST..VW_MASTER_PesertaDtl WHERE ISNULL(Batch,'')<>'' " +
            "  UNION " +
            "  SELECT DISTINCT Batch FROM WISECON_PSIKOTEST.dbo.VW_MASTER_PesertaInterview WHERE ISNULL(Batch,'')<>'' " +
            ") X ORDER BY Batch"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<string>()
            while rdr.Read() do
                let batch = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                if not (String.IsNullOrWhiteSpace(batch)) then rows.Add(batch)
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Applicants/Peserta/List")>]
    member this.ListPeserta([<FromQuery>] inputStart: Nullable<DateTime>, [<FromQuery>] inputEnd: Nullable<DateTime>, [<FromQuery>] posisi: string, [<FromQuery>] hasil: string) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        let mutable baseWhere = " WHERE 1=1 "
        if inputStart.HasValue && inputEnd.HasValue then
            baseWhere <- baseWhere + " AND CAST(P.TimeInput AS DATE) BETWEEN @is AND @ie "
            cmd.Parameters.AddWithValue("@is", inputStart.Value.Date) |> ignore
            cmd.Parameters.AddWithValue("@ie", inputEnd.Value.Date) |> ignore
        match posisi with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s -> baseWhere <- baseWhere + " AND P.LamarSebagai=@pos "
               cmd.Parameters.AddWithValue("@pos", s) |> ignore

        conn.Open()
        try
            try
                let mutable whereClause = baseWhere
                match hasil with
                | null -> ()
                | h when String.IsNullOrWhiteSpace(h) -> ()
                | h when String.Equals(h, "BELUM UJIAN", StringComparison.OrdinalIgnoreCase) -> whereClause <- whereClause + " AND P.LblRek='BELUM UJIAN' "
                | h when String.Equals(h, "LULUS", StringComparison.OrdinalIgnoreCase) -> whereClause <- whereClause + " AND P.LblRek IN ('DISARANKAN','DIPERTIMBANGKAN') "
                | h when String.Equals(h, "TIDAK LULUS", StringComparison.OrdinalIgnoreCase) -> whereClause <- whereClause + " AND P.LblRek='TIDAK DISARANKAN' "
                | _ -> ()

                cmd.CommandText <- "SELECT P.NoPeserta, P.NamaPeserta, P.Alamat, P.Email, P.LamarSebagai, P.NoKTP, P.LblRek, P.TimeInput, P.LastEducation " +
                                   "FROM WISECON_PSIKOTEST.dbo.VW_MASTER_Peserta P" + whereClause + " ORDER BY P.TimeInput DESC"
                use rdr = cmd.ExecuteReader()
                let rows = ResizeArray<obj>()
                let mutable i = 0
                while rdr.Read() do
                    i <- i + 1
                    let nomor = if rdr.IsDBNull(0) then "" else Convert.ToString(rdr.GetValue(0))
                    let nama = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                    let alamat = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                    let email = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                    let posisi = if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                    let ktp = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                    let lblRek = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                    let time = if rdr.IsDBNull(7) then Nullable() else Nullable(rdr.GetDateTime(7))
                    let pendidikan = if rdr.IsDBNull(8) then "" else rdr.GetString(8)
                    let hasil =
                        match lblRek with
                        | null -> ""
                        | s when String.IsNullOrWhiteSpace(s) -> ""
                        | s when String.Equals(s, "DISARANKAN", StringComparison.OrdinalIgnoreCase) -> "LULUS"
                        | s when String.Equals(s, "DIPERTIMBANGKAN", StringComparison.OrdinalIgnoreCase) -> "LULUS"
                        | s when String.Equals(s, "TIDAK DISARANKAN", StringComparison.OrdinalIgnoreCase) -> "TIDAK LULUS"
                        | s -> s
                    rows.Add(box {| num = i; nomor = nomor; nama = nama; alamat = alamat; email = email; pendidikan = pendidikan; posisi = posisi; ktp = ktp; hasil = hasil; time = time |})
                this.Ok(rows)
            with _ ->
                if (not (isNull hasil)) && (not (String.IsNullOrWhiteSpace(hasil))) && (not (String.Equals(hasil, "BELUM UJIAN", StringComparison.OrdinalIgnoreCase))) then
                    this.Ok(Array.empty<obj>)
                else
                    cmd.CommandText <- "SELECT P.NoPeserta, P.NamaPeserta, P.Alamat, P.Email, P.LamarSebagai, P.NoKTP, P.TimeInput, P.PendidikanTerakhir " +
                                       "FROM WISECON_PSIKOTEST.dbo.MS_Peserta P" + baseWhere + " ORDER BY P.TimeInput DESC"
                    use rdr = cmd.ExecuteReader()
                    let rows = ResizeArray<obj>()
                    let mutable i = 0
                    while rdr.Read() do
                        i <- i + 1
                        let nomor = if rdr.IsDBNull(0) then "" else Convert.ToString(rdr.GetValue(0))
                        let nama = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                        let alamat = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                        let email = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                        let posisi = if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                        let ktp = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                        let time = if rdr.IsDBNull(6) then Nullable() else Nullable(rdr.GetDateTime(6))
                        let pendidikan = if rdr.IsDBNull(7) then "" else rdr.GetString(7)
                        rows.Add(box {| num = i; nomor = nomor; nama = nama; alamat = alamat; email = email; pendidikan = pendidikan; posisi = posisi; ktp = ktp; hasil = "BELUM UJIAN"; time = time |})
                    this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Applicants/Ujian/List")>]
    member this.ListUjian([<FromQuery>] inputStart: Nullable<DateTime>, [<FromQuery>] inputEnd: Nullable<DateTime>, [<FromQuery>] ujianStart: Nullable<DateTime>, [<FromQuery>] ujianEnd: Nullable<DateTime>, [<FromQuery>] status: string, [<FromQuery>] hasil: string, [<FromQuery>] noPeserta: string) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        let mutable whereClause = " WHERE 1=1 "
        if inputStart.HasValue && inputEnd.HasValue then
            whereClause <- whereClause + " AND CAST(A.TimeInput AS DATE) BETWEEN @is AND @ie "
            cmd.Parameters.AddWithValue("@is", inputStart.Value.Date) |> ignore
            cmd.Parameters.AddWithValue("@ie", inputEnd.Value.Date) |> ignore
        if ujianStart.HasValue && ujianEnd.HasValue then
            whereClause <- whereClause + " AND CAST(A.WaktuTest AS DATE) BETWEEN @us AND @ue "
            cmd.Parameters.AddWithValue("@us", ujianStart.Value.Date) |> ignore
            cmd.Parameters.AddWithValue("@ue", ujianEnd.Value.Date) |> ignore
        match status with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s when String.Equals(s, "Belum Ujian", StringComparison.OrdinalIgnoreCase) ->
            whereClause <- whereClause + " AND A.StartTest IS NULL "
        | s when String.Equals(s, "Sedang Ujian", StringComparison.OrdinalIgnoreCase) ->
            whereClause <- whereClause + " AND A.StartTest IS NOT NULL AND A.TimeEdit IS NULL "
        | s when String.Equals(s, "Selesai Ujian", StringComparison.OrdinalIgnoreCase) ->
            whereClause <- whereClause + " AND A.StartTest IS NOT NULL AND A.TimeEdit IS NOT NULL "
        | _ -> ()
        match hasil with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s when String.Equals(s, "BELUM UJIAN", StringComparison.OrdinalIgnoreCase) ->
            whereClause <- whereClause + " AND A.StartTest IS NULL "
        | _ -> ()

        match noPeserta with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s ->
            whereClause <- whereClause + " AND (CONVERT(VARCHAR(50),A.NoPeserta)=@noPeserta OR TRY_CONVERT(BIGINT, CONVERT(VARCHAR(50),A.NoPeserta))=TRY_CONVERT(BIGINT, @noPeserta)) "
            cmd.Parameters.AddWithValue("@noPeserta", s.Trim()) |> ignore

        cmd.CommandText <-
            "SELECT " +
            "ISNULL(A.Batch,'') AS Batch, " +
            "A.NoPeserta, " +
            "A.NamaPeserta, " +
            "A.UserId, " +
            "ISNULL(A.UndangPsikotestKe,0) AS UndangPsikotestKe, " +
            "CASE " +
            "WHEN A.StartTest IS NULL THEN 'Belum Ujian' " +
            "WHEN A.TimeEdit IS NULL THEN 'Sedang Ujian' " +
            "ELSE 'Selesai Ujian' END AS StatusPengerjaan, " +
            "ISNULL(P.LblRek,'') AS LblRek, " +
            "ISNULL(PS.NamaPaket,'') AS NamaPaket, " +
            "A.Url, " +
            "A.WaktuTest, " +
            "A.StartTest, " +
            "A.bKirim, " +
            "WA.TimeInput AS WaktuKirim, " +
            "NULL AS WaktuTerbaca, " +
            "A.UserInput, " +
            "A.TimeInput, " +
            "A.UserEdit, " +
            "A.TimeEdit " +
            "FROM WISECON_PSIKOTEST.dbo.VW_MASTER_PesertaDtl A " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.VW_MASTER_Peserta P ON P.NoPeserta=A.NoPeserta " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.MS_PaketSoal PS ON PS.NoPaket=A.NoPaket " +
            "OUTER APPLY (SELECT TOP 1 TimeInput FROM WISECON_PSIKOTEST.dbo.Trx_SendMessageWhatsApp WHERE idclient = P.NoHP ORDER BY Id DESC) WA " +
            whereClause + " ORDER BY A.TimeInput DESC"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            let mutable i = 0
            while rdr.Read() do
                i <- i + 1
                let batch = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                let nomor = if rdr.IsDBNull(1) then "" else Convert.ToString(rdr.GetValue(1))
                let nama = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                let userId = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                let undangKe = if rdr.IsDBNull(4) then 0 else Convert.ToInt32(rdr.GetValue(4))
                let status = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                let lblRek = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                let paket = if rdr.IsDBNull(7) then "" else rdr.GetString(7)
                let link = if rdr.IsDBNull(8) then "" else rdr.GetString(8)
                let waktuTest = if rdr.IsDBNull(9) then Nullable() else Nullable(rdr.GetDateTime(9))
                let waktuMulai = if rdr.IsDBNull(10) then Nullable() else Nullable(rdr.GetDateTime(10))
                let bKirimObj = if rdr.IsDBNull(11) then null else rdr.GetValue(11)
                let statusPesan =
                    match bKirimObj with
                    | null -> ""
                    | :? bool as b -> if b then "Terkirim" else "Belum"
                    | :? int as x -> if x <> 0 then "Terkirim" else "Belum"
                    | :? int16 as x -> if x <> 0s then "Terkirim" else "Belum"
                    | :? int32 as x -> if x <> 0 then "Terkirim" else "Belum"
                    | :? int64 as x -> if x <> 0L then "Terkirim" else "Belum"
                    | _ -> "Belum"
                let kirim = if rdr.IsDBNull(12) then Nullable() else Nullable(rdr.GetDateTime(12))
                let baca = if rdr.IsDBNull(13) then Nullable() else Nullable(rdr.GetDateTime(13))
                let userInput = if rdr.IsDBNull(14) then "" else rdr.GetString(14)
                let timeInput = if rdr.IsDBNull(15) then Nullable() else Nullable(rdr.GetDateTime(15))
                let userEdit = if rdr.IsDBNull(16) then "" else rdr.GetString(16)
                let timeEdit = if rdr.IsDBNull(17) then Nullable() else Nullable(rdr.GetDateTime(17))
                let hasil =
                    match lblRek with
                    | null -> ""
                    | s when String.IsNullOrWhiteSpace(s) -> ""
                    | s when String.Equals(s, "DISARANKAN", StringComparison.OrdinalIgnoreCase) -> "LULUS"
                    | s when String.Equals(s, "DIPERTIMBANGKAN", StringComparison.OrdinalIgnoreCase) -> "LULUS"
                    | s when String.Equals(s, "TIDAK DISARANKAN", StringComparison.OrdinalIgnoreCase) -> "TIDAK LULUS"
                    | s -> s
                rows.Add(box {| num = i; batch = batch; nomor = nomor; nama = nama; userId = userId; undangKe = undangKe; status = status; hasil = hasil; link = link; paket = paket; waktuTest = waktuTest; waktuMulai = waktuMulai; statusPesan = statusPesan; kirim = kirim; baca = baca; userInput = userInput; timeInput = timeInput; userEdit = userEdit; timeEdit = timeEdit |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Applicants/Ujian/Export")>]
    member this.ExportUjian([<FromQuery>] inputStart: Nullable<DateTime>, [<FromQuery>] inputEnd: Nullable<DateTime>, [<FromQuery>] ujianStart: Nullable<DateTime>, [<FromQuery>] ujianEnd: Nullable<DateTime>, [<FromQuery>] status: string, [<FromQuery>] hasil: string, [<FromQuery>] sortField: string, [<FromQuery>] sortDir: Nullable<int>, [<FromQuery>] noPeserta: string) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text

        let mutable whereClause = " WHERE 1=1 "

        let tryGetQuery (key: string) =
            if this.Request.Query.ContainsKey(key) then
                let v = this.Request.Query.[key].ToString()
                if String.IsNullOrWhiteSpace(v) then None else Some v
            else
                None

        let addLike (key: string) (sqlExpr: string) (paramName: string) =
            match tryGetQuery key with
            | None -> ()
            | Some v ->
                whereClause <- whereClause + (" AND " + sqlExpr + " LIKE " + paramName + " ")
                cmd.Parameters.AddWithValue(paramName, "%" + v + "%") |> ignore

        let statusExpr =
            "CASE " +
            "WHEN A.StartTest IS NULL THEN 'Belum Ujian' " +
            "WHEN A.TimeEdit IS NULL THEN 'Sedang Ujian' " +
            "ELSE 'Selesai Ujian' END"

        let statusPesanExpr = "CASE WHEN ISNULL(A.bKirim,0)=1 THEN 'Terkirim' ELSE 'Belum' END"
        if inputStart.HasValue && inputEnd.HasValue then
            whereClause <- whereClause + " AND CAST(A.TimeInput AS DATE) BETWEEN @is AND @ie "
            cmd.Parameters.AddWithValue("@is", inputStart.Value.Date) |> ignore
            cmd.Parameters.AddWithValue("@ie", inputEnd.Value.Date) |> ignore
        if ujianStart.HasValue && ujianEnd.HasValue then
            whereClause <- whereClause + " AND CAST(A.WaktuTest AS DATE) BETWEEN @us AND @ue "
            cmd.Parameters.AddWithValue("@us", ujianStart.Value.Date) |> ignore
            cmd.Parameters.AddWithValue("@ue", ujianEnd.Value.Date) |> ignore
        match status with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s when String.Equals(s, "Belum Ujian", StringComparison.OrdinalIgnoreCase) ->
            whereClause <- whereClause + " AND A.StartTest IS NULL "
        | s when String.Equals(s, "Sedang Ujian", StringComparison.OrdinalIgnoreCase) ->
            whereClause <- whereClause + " AND A.StartTest IS NOT NULL AND A.TimeEdit IS NULL "
        | s when String.Equals(s, "Selesai Ujian", StringComparison.OrdinalIgnoreCase) ->
            whereClause <- whereClause + " AND A.StartTest IS NOT NULL AND A.TimeEdit IS NOT NULL "
        | _ -> ()
        match hasil with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s when String.Equals(s, "BELUM UJIAN", StringComparison.OrdinalIgnoreCase) ->
            whereClause <- whereClause + " AND A.StartTest IS NULL "
        | _ -> ()

        match noPeserta with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s ->
            whereClause <- whereClause + " AND (CONVERT(VARCHAR(50),A.NoPeserta)=@noPeserta OR TRY_CONVERT(BIGINT, CONVERT(VARCHAR(50),A.NoPeserta))=TRY_CONVERT(BIGINT, @noPeserta)) "
            cmd.Parameters.AddWithValue("@noPeserta", s.Trim()) |> ignore

        addLike "q_nomor" "CONVERT(VARCHAR(50),A.NoPeserta)" "@q_nomor"
        addLike "q_nama" "ISNULL(A.NamaPeserta,'')" "@q_nama"
        addLike "q_userId" "ISNULL(A.UserId,'')" "@q_userId"
        addLike "q_link" "ISNULL(A.Url,'')" "@q_link"
        addLike "q_waktuTest" "CONVERT(VARCHAR(19),A.WaktuTest,120)" "@q_waktuTest"
        addLike "q_waktuMulai" "CONVERT(VARCHAR(19),A.StartTest,120)" "@q_waktuMulai"
        addLike "q_statusPesan" statusPesanExpr "@q_statusPesan"
        addLike "q_kirim" "CONVERT(VARCHAR(19),WA.TimeInput,120)" "@q_kirim"
        addLike "q_userInput" "ISNULL(A.UserInput,'')" "@q_userInput"
        addLike "q_timeInput" "CONVERT(VARCHAR(19),A.TimeInput,120)" "@q_timeInput"
        addLike "q_userEdit" "ISNULL(A.UserEdit,'')" "@q_userEdit"
        addLike "q_timeEdit" "CONVERT(VARCHAR(19),A.TimeEdit,120)" "@q_timeEdit"
        addLike "q_status" statusExpr "@q_status"

        let dir = if sortDir.HasValue && sortDir.Value < 0 then "DESC" else "ASC"
        let orderBy =
            match sortField with
            | null -> "A.TimeInput DESC"
            | s when String.IsNullOrWhiteSpace(s) -> "A.TimeInput DESC"
            | s ->
                match s.Trim() with
                | "nomor" -> "A.NoPeserta " + dir
                | "nama" -> "A.NamaPeserta " + dir
                | "userId" -> "A.UserId " + dir
                | "link" -> "A.Url " + dir
                | "waktuTest" -> "A.WaktuTest " + dir
                | "waktuMulai" -> "A.StartTest " + dir
                | "statusPesan" -> "A.bKirim " + dir
                | "kirim" -> "WA.TimeInput " + dir
                | "userInput" -> "A.UserInput " + dir
                | "timeInput" -> "A.TimeInput " + dir
                | "userEdit" -> "A.UserEdit " + dir
                | "timeEdit" -> "A.TimeEdit " + dir
                | _ -> "A.TimeInput DESC"

        cmd.CommandText <-
            "SELECT " +
            "ISNULL(A.Batch,'') AS Batch, " +
            "A.NoPeserta, " +
            "A.NamaPeserta, " +
            "A.UserId, " +
            "ISNULL(A.UndangPsikotestKe,0) AS UndangPsikotestKe, " +
            statusExpr + " AS StatusPengerjaan, " +
            "ISNULL(P.LblRek,'') AS LblRek, " +
            "ISNULL(PS.NamaPaket,'') AS NamaPaket, " +
            "A.Url, " +
            "A.WaktuTest, " +
            "A.StartTest, " +
            "A.bKirim, " +
            "WA.TimeInput AS WaktuKirim, " +
            "NULL AS WaktuTerbaca, " +
            "A.UserInput, " +
            "A.TimeInput, " +
            "A.UserEdit, " +
            "A.TimeEdit " +
            "FROM WISECON_PSIKOTEST.dbo.VW_MASTER_PesertaDtl A " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.VW_MASTER_Peserta P ON P.NoPeserta=A.NoPeserta " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.MS_PaketSoal PS ON PS.NoPaket=A.NoPaket " +
            "OUTER APPLY (SELECT TOP 1 TimeInput FROM WISECON_PSIKOTEST.dbo.Trx_SendMessageWhatsApp WHERE idclient = P.NoHP ORDER BY Id DESC) WA " +
            whereClause + " ORDER BY " + orderBy

        let fmtDt (v: Nullable<DateTime>) =
            if v.HasValue then v.Value.ToString("yyyy-MM-dd HH:mm") else ""

        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let headers =
                [|
                    "Num"; "Batch"; "Nomor Peserta"; "Nama Peserta"; "User ID"; "Undang Ke"; "Status Pengerjaan"; "Hasil Psikotest"; "Link Login"; "Nama Paket Soal"; "Waktu Test"; "Waktu Dimulai"; "Status Pesan"; "Waktu Terkirim"; "Waktu Terbaca"; "User Input"; "Time Input"; "User Edit"; "Time Edit"
                |]
            let data = ResizeArray<string array>()
            let mutable i = 0
            while rdr.Read() do
                i <- i + 1
                let batch = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                let nomor = if rdr.IsDBNull(1) then "" else Convert.ToString(rdr.GetValue(1))
                let nama = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                let userId = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                let undangKe = if rdr.IsDBNull(4) then "" else Convert.ToString(rdr.GetValue(4))
                let statusP = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                let lblRek = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                let paket = if rdr.IsDBNull(7) then "" else rdr.GetString(7)
                let link = if rdr.IsDBNull(8) then "" else rdr.GetString(8)
                let waktuTest = if rdr.IsDBNull(9) then Nullable() else Nullable(rdr.GetDateTime(9))
                let waktuMulai = if rdr.IsDBNull(10) then Nullable() else Nullable(rdr.GetDateTime(10))
                let bKirimObj = if rdr.IsDBNull(11) then null else rdr.GetValue(11)
                let statusPesan =
                    match bKirimObj with
                    | null -> ""
                    | :? bool as b -> if b then "Terkirim" else "Belum"
                    | :? int as x -> if x <> 0 then "Terkirim" else "Belum"
                    | :? int16 as x -> if x <> 0s then "Terkirim" else "Belum"
                    | :? int32 as x -> if x <> 0 then "Terkirim" else "Belum"
                    | :? int64 as x -> if x <> 0L then "Terkirim" else "Belum"
                    | _ -> "Belum"
                let kirim = if rdr.IsDBNull(12) then Nullable() else Nullable(rdr.GetDateTime(12))
                let baca = if rdr.IsDBNull(13) then Nullable() else Nullable(rdr.GetDateTime(13))
                let userInput = if rdr.IsDBNull(14) then "" else rdr.GetString(14)
                let timeInput = if rdr.IsDBNull(15) then Nullable() else Nullable(rdr.GetDateTime(15))
                let userEdit = if rdr.IsDBNull(16) then "" else rdr.GetString(16)
                let timeEdit = if rdr.IsDBNull(17) then Nullable() else Nullable(rdr.GetDateTime(17))
                let hasil =
                    match lblRek with
                    | null -> ""
                    | s when String.IsNullOrWhiteSpace(s) -> ""
                    | s when String.Equals(s, "DISARANKAN", StringComparison.OrdinalIgnoreCase) -> "LULUS"
                    | s when String.Equals(s, "DIPERTIMBANGKAN", StringComparison.OrdinalIgnoreCase) -> "LULUS"
                    | s when String.Equals(s, "TIDAK DISARANKAN", StringComparison.OrdinalIgnoreCase) -> "TIDAK LULUS"
                    | s -> s
                data.Add(
                    [|
                        string i
                        batch
                        nomor
                        nama
                        userId
                        undangKe
                        statusP
                        hasil
                        link
                        paket
                        fmtDt waktuTest
                        fmtDt waktuMulai
                        statusPesan
                        fmtDt kirim
                        fmtDt baca
                        userInput
                        fmtDt timeInput
                        userEdit
                        fmtDt timeEdit
                    |]
                )

            use wb = new XLWorkbook()
            let ws = wb.Worksheets.Add("Psikotest")
            for c = 0 to headers.Length - 1 do
                ws.Cell(1, c + 1).Value <- headers[c]
            let headerRange = ws.Range(1, 1, 1, headers.Length)
            headerRange.Style.Fill.BackgroundColor <- XLColor.FromHtml("#0000FF")
            headerRange.Style.Font.Bold <- true
            headerRange.Style.Font.FontColor <- XLColor.White
            for r = 0 to data.Count - 1 do
                let row = data[r]
                for c = 0 to headers.Length - 1 do
                    ws.Cell(r + 2, c + 1).Value <- row[c]
            ws.Columns().AdjustToContents() |> ignore
            use ms = new MemoryStream()
            wb.SaveAs(ms)
            let bytes = ms.ToArray()
            let fileName = "psikotest.xlsx"
            this.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName) :> IActionResult
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Applicants/Interview/List")>]
    member this.ListInterview([<FromQuery>] inputStart: Nullable<DateTime>, [<FromQuery>] inputEnd: Nullable<DateTime>, [<FromQuery>] statusInterview: string, [<FromQuery>] noPeserta: string) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        let mutable whereClause = " WHERE 1=1 "
        if inputStart.HasValue && inputEnd.HasValue then
            whereClause <- whereClause + " AND CAST(TimeInput AS DATE) BETWEEN @is AND @ie "
            cmd.Parameters.AddWithValue("@is", inputStart.Value.Date) |> ignore
            cmd.Parameters.AddWithValue("@ie", inputEnd.Value.Date) |> ignore
        match statusInterview with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s -> whereClause <- whereClause + " AND StatusInterview=@si "
               cmd.Parameters.AddWithValue("@si", s) |> ignore

        match noPeserta with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s ->
            whereClause <- whereClause + " AND (CONVERT(VARCHAR(50),NoPeserta)=@noPeserta OR TRY_CONVERT(BIGINT, CONVERT(VARCHAR(50),NoPeserta))=TRY_CONVERT(BIGINT, @noPeserta)) "
            cmd.Parameters.AddWithValue("@noPeserta", s.Trim()) |> ignore

        cmd.CommandText <- "SELECT SeqNo, Batch, NoPeserta, NamaPeserta, WaktuInterview, Lokasi, StatusInterview, UndangInterviewKe, bKirim, WaktuKirim, WaktuTerbaca, UserInput, TimeInput, UserEdit, TimeEdit " +
                           "FROM WISECON_PSIKOTEST.dbo.VW_MASTER_PesertaInterview" + whereClause + " ORDER BY TimeInput DESC"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            let mutable i = 0
            while rdr.Read() do
                i <- i + 1
                let seqNo = if rdr.IsDBNull(0) then 0L else Convert.ToInt64(rdr.GetValue(0))
                let batch = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                let nomor = if rdr.IsDBNull(2) then "" else Convert.ToString(rdr.GetValue(2))
                let nama = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                let waktuInterview = if rdr.IsDBNull(4) then Nullable() else Nullable(rdr.GetDateTime(4))
                let lokasi = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                let statusInterview = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                let undangKe = if rdr.IsDBNull(7) then 0 else Convert.ToInt32(rdr.GetValue(7))
                let bKirimObj = if rdr.IsDBNull(8) then null else rdr.GetValue(8)
                let statusPesan =
                    match bKirimObj with
                    | null -> ""
                    | :? bool as b -> if b then "Terkirim" else "Belum"
                    | :? int as x -> if x <> 0 then "Terkirim" else "Belum"
                    | :? int16 as x -> if x <> 0s then "Terkirim" else "Belum"
                    | :? int32 as x -> if x <> 0 then "Terkirim" else "Belum"
                    | :? int64 as x -> if x <> 0L then "Terkirim" else "Belum"
                    | _ -> "Belum"
                let kirim = if rdr.IsDBNull(9) then Nullable() else Nullable(rdr.GetDateTime(9))
                let baca = if rdr.IsDBNull(10) then Nullable() else Nullable(rdr.GetDateTime(10))
                let userInput = if rdr.IsDBNull(11) then "" else rdr.GetString(11)
                let timeInput = if rdr.IsDBNull(12) then Nullable() else Nullable(rdr.GetDateTime(12))
                let userEdit = if rdr.IsDBNull(13) then "" else rdr.GetString(13)
                let timeEdit = if rdr.IsDBNull(14) then Nullable() else Nullable(rdr.GetDateTime(14))
                rows.Add(box {| num = i; seqNo = seqNo; batch = batch; nomor = nomor; nama = nama; waktuInterview = waktuInterview; lokasi = lokasi; statusInterview = statusInterview; undangKe = undangKe; statusPesan = statusPesan; kirim = kirim; baca = baca; userInput = userInput; timeInput = timeInput; userEdit = userEdit; timeEdit = timeEdit |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Applicants/Interview/Export")>]
    member this.ExportInterview([<FromQuery>] inputStart: Nullable<DateTime>, [<FromQuery>] inputEnd: Nullable<DateTime>, [<FromQuery>] statusInterview: string, [<FromQuery>] sortField: string, [<FromQuery>] sortDir: Nullable<int>, [<FromQuery>] noPeserta: string) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text

        let mutable whereClause = " WHERE 1=1 "

        let tryGetQuery (key: string) =
            if this.Request.Query.ContainsKey(key) then
                let v = this.Request.Query.[key].ToString()
                if String.IsNullOrWhiteSpace(v) then None else Some v
            else
                None

        let addLike (key: string) (sqlExpr: string) (paramName: string) =
            match tryGetQuery key with
            | None -> ()
            | Some v ->
                whereClause <- whereClause + (" AND " + sqlExpr + " LIKE " + paramName + " ")
                cmd.Parameters.AddWithValue(paramName, "%" + v + "%") |> ignore

        let statusPesanExpr = "CASE WHEN ISNULL(I.bKirim,0)=1 THEN 'Terkirim' ELSE 'Belum' END"
        if inputStart.HasValue && inputEnd.HasValue then
            whereClause <- whereClause + " AND CAST(I.TimeInput AS DATE) BETWEEN @is AND @ie "
            cmd.Parameters.AddWithValue("@is", inputStart.Value.Date) |> ignore
            cmd.Parameters.AddWithValue("@ie", inputEnd.Value.Date) |> ignore
        match statusInterview with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s -> whereClause <- whereClause + " AND I.StatusInterview=@si "
               cmd.Parameters.AddWithValue("@si", s) |> ignore

        match noPeserta with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s ->
            whereClause <- whereClause + " AND (CONVERT(VARCHAR(50),I.NoPeserta)=@noPeserta OR TRY_CONVERT(BIGINT, CONVERT(VARCHAR(50),I.NoPeserta))=TRY_CONVERT(BIGINT, @noPeserta)) "
            cmd.Parameters.AddWithValue("@noPeserta", s.Trim()) |> ignore

        addLike "q_batch" "ISNULL(I.Batch,'')" "@q_batch"
        addLike "q_nomor" "CONVERT(VARCHAR(50),I.NoPeserta)" "@q_nomor"
        addLike "q_nama" "ISNULL(I.NamaPeserta,'')" "@q_nama"
        addLike "q_waktuInterview" "CONVERT(VARCHAR(19),I.WaktuInterview,120)" "@q_waktuInterview"
        addLike "q_lokasi" "ISNULL(I.Lokasi,'')" "@q_lokasi"
        addLike "q_statusInterview" "ISNULL(I.StatusInterview,'')" "@q_statusInterview"
        addLike "q_undangKe" "CONVERT(VARCHAR(50),I.UndangInterviewKe)" "@q_undangKe"
        addLike "q_statusPesan" statusPesanExpr "@q_statusPesan"
        addLike "q_kirim" "CONVERT(VARCHAR(19),I.WaktuKirim,120)" "@q_kirim"
        addLike "q_baca" "CONVERT(VARCHAR(19),I.WaktuTerbaca,120)" "@q_baca"
        addLike "q_userInput" "ISNULL(I.UserInput,'')" "@q_userInput"
        addLike "q_timeInput" "CONVERT(VARCHAR(19),I.TimeInput,120)" "@q_timeInput"
        addLike "q_userEdit" "ISNULL(I.UserEdit,'')" "@q_userEdit"
        addLike "q_timeEdit" "CONVERT(VARCHAR(19),I.TimeEdit,120)" "@q_timeEdit"

        let dir = if sortDir.HasValue && sortDir.Value < 0 then "DESC" else "ASC"
        let orderBy =
            match sortField with
            | null -> "I.TimeInput DESC"
            | s when String.IsNullOrWhiteSpace(s) -> "I.TimeInput DESC"
            | s ->
                match s.Trim() with
                | "batch" -> "I.Batch " + dir
                | "nomor" -> "I.NoPeserta " + dir
                | "nama" -> "I.NamaPeserta " + dir
                | "waktuInterview" -> "I.WaktuInterview " + dir
                | "lokasi" -> "I.Lokasi " + dir
                | "statusInterview" -> "I.StatusInterview " + dir
                | "undangKe" -> "I.UndangInterviewKe " + dir
                | "statusPesan" -> "I.bKirim " + dir
                | "kirim" -> "I.WaktuKirim " + dir
                | "baca" -> "I.WaktuTerbaca " + dir
                | "userInput" -> "I.UserInput " + dir
                | "timeInput" -> "I.TimeInput " + dir
                | "userEdit" -> "I.UserEdit " + dir
                | "timeEdit" -> "I.TimeEdit " + dir
                | _ -> "I.TimeInput DESC"

        cmd.CommandText <-
            "SELECT I.Batch, I.NoPeserta, I.NamaPeserta, I.WaktuInterview, I.Lokasi, I.StatusInterview, I.UndangInterviewKe, I.bKirim, I.WaktuKirim, I.WaktuTerbaca, I.UserInput, I.TimeInput, I.UserEdit, I.TimeEdit " +
            "FROM WISECON_PSIKOTEST.dbo.VW_MASTER_PesertaInterview I" + whereClause + " ORDER BY " + orderBy

        let fmtDt (v: Nullable<DateTime>) =
            if v.HasValue then v.Value.ToString("yyyy-MM-dd HH:mm") else ""

        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let headers =
                [|
                    "Num"; "Batch"; "Nomor Peserta"; "Nama Peserta"; "Waktu Interview"; "Lokasi"; "Status Interview"; "Undang Ke"; "Status Pesan"; "Waktu Terkirim"; "Waktu Terbaca"; "User Input"; "Time Input"; "User Edit"; "Time Edit"
                |]
            let data = ResizeArray<string array>()
            let mutable i = 0
            while rdr.Read() do
                i <- i + 1
                let batch = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                let nomor = if rdr.IsDBNull(1) then "" else Convert.ToString(rdr.GetValue(1))
                let nama = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                let waktuInterview = if rdr.IsDBNull(3) then Nullable() else Nullable(rdr.GetDateTime(3))
                let lokasi = if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                let statusInterview = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                let undangKe = if rdr.IsDBNull(6) then "" else Convert.ToString(rdr.GetValue(6))
                let bKirimObj = if rdr.IsDBNull(7) then null else rdr.GetValue(7)
                let statusPesan =
                    match bKirimObj with
                    | null -> ""
                    | :? bool as b -> if b then "Terkirim" else "Belum"
                    | :? int as x -> if x <> 0 then "Terkirim" else "Belum"
                    | :? int16 as x -> if x <> 0s then "Terkirim" else "Belum"
                    | :? int32 as x -> if x <> 0 then "Terkirim" else "Belum"
                    | :? int64 as x -> if x <> 0L then "Terkirim" else "Belum"
                    | _ -> "Belum"
                let kirim = if rdr.IsDBNull(8) then Nullable() else Nullable(rdr.GetDateTime(8))
                let baca = if rdr.IsDBNull(9) then Nullable() else Nullable(rdr.GetDateTime(9))
                let userInput = if rdr.IsDBNull(10) then "" else rdr.GetString(10)
                let timeInput = if rdr.IsDBNull(11) then Nullable() else Nullable(rdr.GetDateTime(11))
                let userEdit = if rdr.IsDBNull(12) then "" else rdr.GetString(12)
                let timeEdit = if rdr.IsDBNull(13) then Nullable() else Nullable(rdr.GetDateTime(13))
                data.Add(
                    [|
                        string i
                        batch
                        nomor
                        nama
                        fmtDt waktuInterview
                        lokasi
                        statusInterview
                        undangKe
                        statusPesan
                        fmtDt kirim
                        fmtDt baca
                        userInput
                        fmtDt timeInput
                        userEdit
                        fmtDt timeEdit
                    |]
                )

            use wb = new XLWorkbook()
            let ws = wb.Worksheets.Add("Interview")
            for c = 0 to headers.Length - 1 do
                ws.Cell(1, c + 1).Value <- headers[c]
            let headerRange = ws.Range(1, 1, 1, headers.Length)
            headerRange.Style.Fill.BackgroundColor <- XLColor.FromHtml("#FFA500")
            headerRange.Style.Font.Bold <- true
            headerRange.Style.Font.FontColor <- XLColor.White
            for r = 0 to data.Count - 1 do
                let row = data[r]
                for c = 0 to headers.Length - 1 do
                    ws.Cell(r + 2, c + 1).Value <- row[c]
            ws.Columns().AdjustToContents() |> ignore
            use ms = new MemoryStream()
            wb.SaveAs(ms)
            let bytes = ms.ToArray()
            let fileName = "interview.xlsx"
            this.File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName) :> IActionResult
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Applicants/Ujian/Assign")>]
    member this.AssignUjian([<FromBody>] req: AssignUjianRequest) : IActionResult =
        let domainName = getDomainName ()
        let encryptionKey = getEncryptionKey ()
        if String.IsNullOrWhiteSpace(domainName) || String.IsNullOrWhiteSpace(encryptionKey) then
            this.BadRequest("Konfigurasi DomainName/EncryptionKey belum diset")
        else
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        conn.Open()
        try
            use cmd1 = new Microsoft.Data.SqlClient.SqlCommand()
            cmd1.Connection <- conn
            cmd1.CommandType <- CommandType.StoredProcedure
            cmd1.CommandText <- "WISECON_PSIKOTEST..SP_PesertaDtl"
            cmd1.Parameters.AddWithValue("@User", req.User) |> ignore
            cmd1.Parameters.AddWithValue("@Act", "ADD") |> ignore
            cmd1.Parameters.AddWithValue("@lvl", 1s) |> ignore
            let mutable userId = ""
            let mutable passwordEnc = ""
            let mutable url = ""
            let mutable noKtp = ""
            let mutable email = ""
            use cmdPeserta = new Microsoft.Data.SqlClient.SqlCommand()
            cmdPeserta.Connection <- conn
            cmdPeserta.CommandType <- CommandType.Text
            cmdPeserta.CommandText <- "SELECT NoKTP, Email FROM WISECON_PSIKOTEST.dbo.MS_Peserta WHERE ID=@id"
            cmdPeserta.Parameters.AddWithValue("@id", req.NoPeserta) |> ignore
            use rdPeserta = cmdPeserta.ExecuteReader()
            if rdPeserta.Read() then
                noKtp <- if rdPeserta.IsDBNull(0) then "" else rdPeserta.GetString(0)
                email <- if rdPeserta.IsDBNull(1) then "" else rdPeserta.GetString(1)
            rdPeserta.Close()

            use rd = cmd1.ExecuteReader()
            if rd.Read() then
                userId <- rd.GetString(rd.GetOrdinal("NewUserID"))
                let pwd = rd.GetString(rd.GetOrdinal("NewPassword"))
                passwordEnc <- encrypt64 pwd encryptionKey
                let token = "KTP=" + noKtp + "&EM=" + email + "&U=" + userId + "&NP=" + req.NoPaket.ToString() + "&WT=" + req.WaktuUjian.ToString("o")
                let tokenEnc = encrypt64 token encryptionKey |> Uri.EscapeDataString
                let baseUrl = if domainName.EndsWith("/", StringComparison.Ordinal) then domainName else domainName + "/"
                url <- baseUrl + "Exam?Token=" + tokenEnc
            rd.Close()

            use cmd2 = new Microsoft.Data.SqlClient.SqlCommand()
            cmd2.Connection <- conn
            cmd2.CommandType <- CommandType.StoredProcedure
            cmd2.CommandText <- "WISECON_PSIKOTEST..SP_PesertaDtl"
            cmd2.Parameters.AddWithValue("@Batch", (if isNull req.Batch then "" else req.Batch)) |> ignore
            cmd2.Parameters.AddWithValue("@NoPeserta", req.NoPeserta) |> ignore
            cmd2.Parameters.AddWithValue("@NoPaket", req.NoPaket) |> ignore
            cmd2.Parameters.AddWithValue("@UserID", userId) |> ignore
            cmd2.Parameters.AddWithValue("@Password", passwordEnc) |> ignore
            cmd2.Parameters.AddWithValue("@Url", url) |> ignore
            cmd2.Parameters.AddWithValue("@WaktuUjian", req.WaktuUjian) |> ignore
            cmd2.Parameters.AddWithValue("@User", req.User) |> ignore
            cmd2.Parameters.AddWithValue("@Act", "ADD") |> ignore
            cmd2.Parameters.AddWithValue("@lvl", 2s) |> ignore
            use rd2 = cmd2.ExecuteReader()
            let assigned =
                if rd2.Read() then true else false
            rd2.Close()
            use cmdMsg = new Microsoft.Data.SqlClient.SqlCommand()
            cmdMsg.Connection <- conn
            cmdMsg.CommandType <- CommandType.StoredProcedure
            cmdMsg.CommandText <- "dbo.Insert_MessageWhatsapp_Psikotest"
            cmdMsg.Parameters.AddWithValue("@NoPeserta", req.NoPeserta) |> ignore
            cmdMsg.ExecuteNonQuery() |> ignore
            this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Applicants/Ujian/AssignBulk")>]
    member this.AssignUjianBulk([<FromBody>] req: AssignUjianBulkRequest) : IActionResult =
        let domainName = getDomainName ()
        let encryptionKey = getEncryptionKey ()
        if String.IsNullOrWhiteSpace(domainName) || String.IsNullOrWhiteSpace(encryptionKey) then
            this.BadRequest("Konfigurasi DomainName/EncryptionKey belum diset")
        else
            let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
            conn.Open()
            try
                let batch = if isNull req.Batch then "" else req.Batch
                let user = if isNull req.User then "" else req.User
                let noPesertaList = if isNull req.NoPeserta then [||] else req.NoPeserta
                for noPeserta in noPesertaList do
                    use cmd1 = new Microsoft.Data.SqlClient.SqlCommand()
                    cmd1.Connection <- conn
                    cmd1.CommandType <- CommandType.StoredProcedure
                    cmd1.CommandText <- "WISECON_PSIKOTEST..SP_PesertaDtl"
                    cmd1.Parameters.AddWithValue("@User", user) |> ignore
                    cmd1.Parameters.AddWithValue("@Act", "ADD") |> ignore
                    cmd1.Parameters.AddWithValue("@lvl", 1s) |> ignore
                    let mutable userId = ""
                    let mutable passwordEnc = ""
                    let mutable url = ""
                    let mutable noKtp = ""
                    let mutable email = ""
                    use cmdPeserta = new Microsoft.Data.SqlClient.SqlCommand()
                    cmdPeserta.Connection <- conn
                    cmdPeserta.CommandType <- CommandType.Text
                    cmdPeserta.CommandText <- "SELECT NoKTP, Email FROM WISECON_PSIKOTEST.dbo.MS_Peserta WHERE ID=@id"
                    cmdPeserta.Parameters.AddWithValue("@id", noPeserta) |> ignore
                    use rdPeserta = cmdPeserta.ExecuteReader()
                    if rdPeserta.Read() then
                        noKtp <- if rdPeserta.IsDBNull(0) then "" else rdPeserta.GetString(0)
                        email <- if rdPeserta.IsDBNull(1) then "" else rdPeserta.GetString(1)
                    rdPeserta.Close()

                    use rd = cmd1.ExecuteReader()
                    if rd.Read() then
                        userId <- rd.GetString(rd.GetOrdinal("NewUserID"))
                        let pwd = rd.GetString(rd.GetOrdinal("NewPassword"))
                        passwordEnc <- encrypt64 pwd encryptionKey
                        let token = "KTP=" + noKtp + "&EM=" + email + "&U=" + userId + "&NP=" + req.NoPaket.ToString() + "&WT=" + req.WaktuUjian.ToString("o")
                        let tokenEnc = encrypt64 token encryptionKey |> Uri.EscapeDataString
                        let baseUrl = if domainName.EndsWith("/", StringComparison.Ordinal) then domainName else domainName + "/"
                        url <- baseUrl + "Exam?Token=" + tokenEnc
                    rd.Close()

                    use cmd2 = new Microsoft.Data.SqlClient.SqlCommand()
                    cmd2.Connection <- conn
                    cmd2.CommandType <- CommandType.StoredProcedure
                    cmd2.CommandText <- "WISECON_PSIKOTEST..SP_PesertaDtl"
                    cmd2.Parameters.AddWithValue("@Batch", batch) |> ignore
                    cmd2.Parameters.AddWithValue("@NoPeserta", noPeserta) |> ignore
                    cmd2.Parameters.AddWithValue("@NoPaket", req.NoPaket) |> ignore
                    cmd2.Parameters.AddWithValue("@UserID", userId) |> ignore
                    cmd2.Parameters.AddWithValue("@Password", passwordEnc) |> ignore
                    cmd2.Parameters.AddWithValue("@Url", url) |> ignore
                    cmd2.Parameters.AddWithValue("@WaktuUjian", req.WaktuUjian) |> ignore
                    cmd2.Parameters.AddWithValue("@User", user) |> ignore
                    cmd2.Parameters.AddWithValue("@Act", "ADD") |> ignore
                    cmd2.Parameters.AddWithValue("@lvl", 2s) |> ignore
                    use rd2 = cmd2.ExecuteReader()
                    let assigned =
                        if rd2.Read() then true else false
                    rd2.Close()
                    use cmdMsg = new Microsoft.Data.SqlClient.SqlCommand()
                    cmdMsg.Connection <- conn
                    cmdMsg.CommandType <- CommandType.StoredProcedure
                    cmdMsg.CommandText <- "dbo.Insert_MessageWhatsapp_Psikotest"
                    cmdMsg.Parameters.AddWithValue("@NoPeserta", noPeserta) |> ignore
                    cmdMsg.ExecuteNonQuery() |> ignore
                this.Ok()
            finally
                conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Applicants/Ujian/Edit")>]
    member this.EditUjian([<FromBody>] req: EditUjianRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.StoredProcedure
        cmd.CommandText <- "WISECON_PSIKOTEST..SP_PesertaDtl"
        cmd.Parameters.AddWithValue("@NoPaket", req.NoPaket) |> ignore
        cmd.Parameters.AddWithValue("@UserID", req.UserID) |> ignore
        cmd.Parameters.AddWithValue("@WaktuUjian", req.WaktuUjian) |> ignore
        cmd.Parameters.AddWithValue("@block", req.block) |> ignore
        cmd.Parameters.AddWithValue("@User", req.User) |> ignore
        cmd.Parameters.AddWithValue("@Act", "ADD") |> ignore
        cmd.Parameters.AddWithValue("@lvl", 2s) |> ignore
        conn.Open()
        try
            let _ = cmd.ExecuteNonQuery()
            this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Applicants/Ujian/Resend")>]
    member this.ResendUjian([<FromBody>] req: ResendUjianRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.StoredProcedure
        cmd.CommandText <- "WISECON_PSIKOTEST..SP_PesertaDtl"
        cmd.Parameters.AddWithValue("@UserID", req.UserID) |> ignore
        cmd.Parameters.AddWithValue("@User", req.User) |> ignore
        cmd.Parameters.AddWithValue("@Act", "ADD") |> ignore
        cmd.Parameters.AddWithValue("@lvl", 3s) |> ignore
        conn.Open()
        try
            let _ = cmd.ExecuteNonQuery()
            this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Applicants/Interview/Assign")>]
    member this.AssignInterview([<FromBody>] req: AssignInterviewRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        conn.Open()
        try
            use cmd = new Microsoft.Data.SqlClient.SqlCommand()
            cmd.Connection <- conn
            cmd.CommandType <- CommandType.StoredProcedure
            cmd.CommandText <- "WISECON_PSIKOTEST..SP_PesertaInterview"
            cmd.Parameters.AddWithValue("@Batch", (if isNull req.Batch then "" else req.Batch)) |> ignore
            cmd.Parameters.AddWithValue("@NoPeserta", req.NoPeserta) |> ignore
            cmd.Parameters.AddWithValue("@WaktuInterview", req.WaktuInterview) |> ignore
            cmd.Parameters.AddWithValue("@Lokasi", (if isNull req.Lokasi then "" else req.Lokasi)) |> ignore
            cmd.Parameters.AddWithValue("@User", req.User) |> ignore
            cmd.Parameters.AddWithValue("@Act", "ADD") |> ignore
            cmd.Parameters.AddWithValue("@lvl", 1s) |> ignore
            use rdr = cmd.ExecuteReader()
            let mutable id = 0L
            if rdr.Read() then
                let ord = rdr.GetOrdinal("ID")
                id <- if rdr.IsDBNull(ord) then 0L else Convert.ToInt64(rdr.GetValue(ord))
            rdr.Close()
            use cmdMsg = new Microsoft.Data.SqlClient.SqlCommand()
            cmdMsg.Connection <- conn
            cmdMsg.CommandType <- CommandType.StoredProcedure
            cmdMsg.CommandText <- "dbo.Insert_MessageWhatsapp_Interview"
            cmdMsg.Parameters.AddWithValue("@NoPeserta", req.NoPeserta) |> ignore
            cmdMsg.ExecuteNonQuery() |> ignore
            this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Applicants/Interview/AssignBulk")>]
    member this.AssignInterviewBulk([<FromBody>] req: AssignInterviewBulkRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        conn.Open()
        try
            let batch = if isNull req.Batch then "" else req.Batch
            let user = if isNull req.User then "" else req.User
            let lokasi = if isNull req.Lokasi then "" else req.Lokasi
            let noPesertaList = if isNull req.NoPeserta then [||] else req.NoPeserta
            for noPeserta in noPesertaList do
                use cmd = new Microsoft.Data.SqlClient.SqlCommand()
                cmd.Connection <- conn
                cmd.CommandType <- CommandType.StoredProcedure
                cmd.CommandText <- "WISECON_PSIKOTEST..SP_PesertaInterview"
                cmd.Parameters.AddWithValue("@Batch", batch) |> ignore
                cmd.Parameters.AddWithValue("@NoPeserta", noPeserta) |> ignore
                cmd.Parameters.AddWithValue("@WaktuInterview", req.WaktuInterview) |> ignore
                cmd.Parameters.AddWithValue("@Lokasi", lokasi) |> ignore
                cmd.Parameters.AddWithValue("@User", user) |> ignore
                cmd.Parameters.AddWithValue("@Act", "ADD") |> ignore
                cmd.Parameters.AddWithValue("@lvl", 1s) |> ignore
                use rdr = cmd.ExecuteReader()
                let mutable id = 0L
                if rdr.Read() then
                    let ord = rdr.GetOrdinal("ID")
                    id <- if rdr.IsDBNull(ord) then 0L else Convert.ToInt64(rdr.GetValue(ord))
                rdr.Close()
                use cmdMsg = new Microsoft.Data.SqlClient.SqlCommand()
                cmdMsg.Connection <- conn
                cmdMsg.CommandType <- CommandType.StoredProcedure
                cmdMsg.CommandText <- "dbo.Insert_MessageWhatsapp_Interview"
                cmdMsg.Parameters.AddWithValue("@NoPeserta", noPeserta) |> ignore
                cmdMsg.ExecuteNonQuery() |> ignore
            this.Ok()
        finally
            conn.Close()
    [<Authorize>]
    [<HttpPost>]
    [<Route("Applicants/Interview/Edit")>]
    member this.EditInterview([<FromBody>] req: EditInterviewRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.StoredProcedure
        cmd.CommandText <- "WISECON_PSIKOTEST..SP_PesertaInterview"
        cmd.Parameters.AddWithValue("@ID", req.SeqNo) |> ignore
        cmd.Parameters.AddWithValue("@Batch", (if isNull req.Batch then "" else req.Batch)) |> ignore
        cmd.Parameters.AddWithValue("@WaktuInterview", req.WaktuInterview) |> ignore
        cmd.Parameters.AddWithValue("@Lokasi", (if isNull req.Lokasi then "" else req.Lokasi)) |> ignore
        cmd.Parameters.AddWithValue("@User", req.User) |> ignore
        cmd.Parameters.AddWithValue("@Act", "ADD") |> ignore
        cmd.Parameters.AddWithValue("@lvl", 1s) |> ignore
        conn.Open()
        try
            let _ = cmd.ExecuteNonQuery()
            this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Applicants/Interview/Resend")>]
    member this.ResendInterview([<FromBody>] req: ResendInterviewRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.StoredProcedure
        cmd.CommandText <- "WISECON_PSIKOTEST..SP_PesertaInterview"
        cmd.Parameters.AddWithValue("@ID", req.SeqNo) |> ignore
        cmd.Parameters.AddWithValue("@User", req.User) |> ignore
        cmd.Parameters.AddWithValue("@Act", "ADD") |> ignore
        cmd.Parameters.AddWithValue("@lvl", 2s) |> ignore
        conn.Open()
        try
            let _ = cmd.ExecuteNonQuery()
            this.Ok()
        finally
            conn.Close()
