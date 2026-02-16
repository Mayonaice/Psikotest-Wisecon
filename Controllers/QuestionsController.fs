namespace PsikotestWisesa.Controllers

open System
open System.Data
open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authorization
open Microsoft.AspNetCore.Http
open Microsoft.AspNetCore.Hosting

[<CLIMutable>]
type PaketRequest = {
    NoPaket: Nullable<int64>
    NamaPaket: string
    ToleransiWaktu: Nullable<int>
    Status: string
    Position: string
    User: string
}

[<CLIMutable>]
type GroupRequest = {
    NoPaket: int64
    NoGroup: Nullable<int64>
    NamaGroup: string
    MinimumJmlSoal: Nullable<int>
    NilaiStandar: Nullable<int>
    WaktuPengerjaan: Nullable<int>
    NoPetunjuk: Nullable<int>
    bRandom: bool
    IsPrioritas: bool
    bAktif: bool
    User: string
}

[<CLIMutable>]
type DeleteGroupRequest = { NoGroup: int64 }

[<CLIMutable>]
type DtlRequest = {
    SeqNo: Nullable<int64>
    NoPaket: int64
    NoGroup: int64
    NoUrut: int64
    Judul: string
    Deskripsi: string
    IsDownload: bool
    bAktif: bool
    MediaFileName: string
    UrlMedia: string
    TipeMedia: string
    User: string
}

[<CLIMutable>]
type DeleteDtlRequest = { SeqNo: int64 }

[<CLIMutable>]
type CopyDtlRequest = {
    SeqNo: int64
    DestNoGroup: int64
    User: string
}

[<CLIMutable>]
type JawabanRequest = {
    NoPaket: int64
    NoGroup: int64
    NoUrut: int64
    NoJawaban: int
    Jawaban: string
    PoinJawaban: int
    MediaFileName: string
    UrlMedia: string
    TipeMedia: string
    TextMedia: string
    User: string
}

[<CLIMutable>]
type DeleteJawabanRequest = { SeqNo: int64 }

[<CLIMutable>]
type NormaRequest = {
    SeqNo: Nullable<int64>
    NoGroup: int64
    LabelNorma: string
    BatasAtas: Nullable<int>
    BatasBawah: Nullable<int>
    User: string
}

[<CLIMutable>]
type DeleteNormaRequest = { SeqNo: int64 }

[<ApiController>]
type QuestionsController (db: System.Data.IDbConnection) =
    inherit Controller()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Paket/Add")>]
    member this.AddPaket ([<FromBody>] req: PaketRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.Parameters.AddWithValue("@NamaPaket", (if isNull req.NamaPaket then "" else req.NamaPaket)) |> ignore
        cmd.Parameters.AddWithValue("@ToleransiWaktu", (if req.ToleransiWaktu.HasValue then req.ToleransiWaktu.Value else 0)) |> ignore
        cmd.Parameters.AddWithValue("@bAktif", (if String.Equals(req.Status, "AKTIF", StringComparison.OrdinalIgnoreCase) then 1 else 0)) |> ignore
        cmd.Parameters.AddWithValue("@User", (if isNull req.User then "" else req.User)) |> ignore
        cmd.Parameters.AddWithValue("@Position", (if isNull req.Position then "" else req.Position)) |> ignore
        conn.Open()
        try
            use chk = new Microsoft.Data.SqlClient.SqlCommand()
            chk.Connection <- conn
            chk.CommandType <- CommandType.Text
            chk.CommandText <- "SELECT CASE WHEN EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('WISECON_PSIKOTEST.dbo.MS_PaketSoal') AND name = 'Position') THEN 1 ELSE 0 END"
            let vObj = chk.ExecuteScalar()
            let hasPos = try Convert.ToInt32(vObj) = 1 with _ -> false
            cmd.CommandText <- (if hasPos then
                                    "INSERT INTO WISECON_PSIKOTEST.dbo.MS_PaketSoal (NamaPaket, ToleransiWaktu, Position, bAktif, UserInput, TimeInput, UserEdit, TimeEdit) VALUES (@NamaPaket, @ToleransiWaktu, @Position, @bAktif, @User, GETDATE(), @User, GETDATE()); SELECT SCOPE_IDENTITY()"
                                else
                                    "INSERT INTO WISECON_PSIKOTEST.dbo.MS_PaketSoal (NamaPaket, ToleransiWaktu, Position, bAktif, UserInput, TimeInput, UserEdit, TimeEdit) VALUES (@NamaPaket, @ToleransiWaktu, @bAktif, @User, GETDATE(), @User, GETDATE()); SELECT SCOPE_IDENTITY()")
            let idObj = cmd.ExecuteScalar()
            let id = (if isNull idObj then 0L else Convert.ToInt64(idObj))
            this.Ok(box {| NoPaket = id |})
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Paket/Edit")>]
    member this.EditPaket ([<FromBody>] req: PaketRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.Parameters.AddWithValue("@NoPaket", (if req.NoPaket.HasValue then req.NoPaket.Value else 0L)) |> ignore
        cmd.Parameters.AddWithValue("@NamaPaket", (if isNull req.NamaPaket then "" else req.NamaPaket)) |> ignore
        cmd.Parameters.AddWithValue("@ToleransiWaktu", (if req.ToleransiWaktu.HasValue then req.ToleransiWaktu.Value else 0)) |> ignore
        let bAktif = if String.Equals(req.Status, "AKTIF", StringComparison.OrdinalIgnoreCase) then 1 else 0
        cmd.Parameters.AddWithValue("@bAktif", bAktif) |> ignore
        cmd.Parameters.AddWithValue("@User", (if isNull req.User then "" else req.User)) |> ignore
        cmd.Parameters.AddWithValue("@Position", (if isNull req.Position then "" else req.Position)) |> ignore
        conn.Open()
        try
            use chk = new Microsoft.Data.SqlClient.SqlCommand()
            chk.Connection <- conn
            chk.CommandType <- CommandType.Text
            chk.CommandText <- "SELECT CASE WHEN EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('WISECON_PSIKOTEST.dbo.MS_PaketSoal') AND name = 'Position') THEN 1 ELSE 0 END"
            let vObj = chk.ExecuteScalar()
            let hasPos = try Convert.ToInt32(vObj) = 1 with _ -> false
            cmd.CommandText <- (if hasPos then
                                    "UPDATE WISECON_PSIKOTEST.dbo.MS_PaketSoal SET NamaPaket=@NamaPaket, ToleransiWaktu=@ToleransiWaktu, Position=@Position, bAktif=@bAktif, UserEdit=@User, TimeEdit=GETDATE() WHERE NoPaket=@NoPaket"
                                else   
                                    "UPDATE WISECON_PSIKOTEST.dbo.MS_PaketSoal SET NamaPaket=@NamaPaket, ToleransiWaktu=@ToleransiWaktu, bAktif=@bAktif, UserEdit=@User, TimeEdit=GETDATE() WHERE NoPaket=@NoPaket")
            if bAktif = 1 then
                use chk = new Microsoft.Data.SqlClient.SqlCommand()
                chk.Connection <- conn
                chk.CommandText <- "SELECT COUNT(*) FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup WHERE NoPaket=@NoPaket AND ISNULL(bAktif,0)=1"
                chk.CommandType <- CommandType.Text
                chk.Parameters.AddWithValue("@NoPaket", (if req.NoPaket.HasValue then req.NoPaket.Value else 0L)) |> ignore
                let cObj = chk.ExecuteScalar()
                let cnt = if isNull cObj then 0 else (try Convert.ToInt32(cObj) with _ -> 0)
                if cnt = 0 then
                    this.BadRequest(box {| error = "Paket Soal harus memiliki minimal 1 Group aktif sebelum diaktifkan" |})
                else
                    let _ = cmd.ExecuteNonQuery()
                    this.Ok()
            else
                let _ = cmd.ExecuteNonQuery()
                this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Group/Modify")>]
    member this.ModifyGroup ([<FromBody>] req: GroupRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandText <- "WISECON_PSIKOTEST.dbo.SP_PaketSoalGroup"
        cmd.CommandType <- CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@NoPaket", req.NoPaket) |> ignore
        cmd.Parameters.AddWithValue("@NoGroup", (if req.NoGroup.HasValue then req.NoGroup.Value else 0L)) |> ignore
        cmd.Parameters.AddWithValue("@NamaGroup", (if isNull req.NamaGroup then "" else req.NamaGroup)) |> ignore
        cmd.Parameters.AddWithValue("@MinimumJmlSoal", (if req.MinimumJmlSoal.HasValue then req.MinimumJmlSoal.Value else 0)) |> ignore
        cmd.Parameters.AddWithValue("@NilaiStandar", (if req.NilaiStandar.HasValue then req.NilaiStandar.Value else 0)) |> ignore
        cmd.Parameters.AddWithValue("@WaktuPengerjaan", (if req.WaktuPengerjaan.HasValue then req.WaktuPengerjaan.Value else 0)) |> ignore
        cmd.Parameters.AddWithValue("@NoPetunjuk", (if req.NoPetunjuk.HasValue then req.NoPetunjuk.Value else 0)) |> ignore
        cmd.Parameters.AddWithValue("@bRandom", req.bRandom) |> ignore
        cmd.Parameters.AddWithValue("@IsPrioritas", req.IsPrioritas) |> ignore
        cmd.Parameters.AddWithValue("@bAktif", req.bAktif) |> ignore
        cmd.Parameters.AddWithValue("@User", (if isNull req.User then "" else req.User)) |> ignore
        let act = if req.NoGroup.HasValue && req.NoGroup.Value > 0L then "EDIT" else "ADD"
        cmd.Parameters.AddWithValue("@Act", act) |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let mutable noGroup = 0L
            if rdr.Read() then
                try
                    noGroup <- rdr.GetInt64(0)
                with _ -> ()
            rdr.Close()
            this.Ok(box {| NoGroup = noGroup |})
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Questions/Jobs/List")>]
    member this.ListJobs () : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT Code, Name FROM WISECON_PSIKOTEST.dbo.VW_ADVHR__REC_Job ORDER BY Name"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let code = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                let name = if rdr.IsDBNull(1) then code else rdr.GetString(1)
                rows.Add(box {| code = code; name = name |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Questions/Paket/Positions/List")>]
    member this.ListPaketPositions ([<FromQuery>] noPaket: int64) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT CodeJob FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalChild WHERE NoPaket=@p"
        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let code = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                rows.Add(box {| code = code |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Questions/Paket/Posisi/Options")>]
    member this.ListPosisiOptions () : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT DISTINCT LTRIM(RTRIM(LamarSebagai)) AS Posisi FROM WISECON_PSIKOTEST.dbo.VW_MASTER_Peserta WHERE LamarSebagai IS NOT NULL AND LTRIM(RTRIM(LamarSebagai)) <> '' ORDER BY LTRIM(RTRIM(LamarSebagai))"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<string>()
            while rdr.Read() do
                let posisi = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                if not (String.IsNullOrWhiteSpace(posisi)) then rows.Add(posisi)
            this.Ok(rows)
        finally
            conn.Close()
    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Group/Delete")>]
    member this.DeleteGroup ([<FromBody>] req: DeleteGroupRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        conn.Open()
        try
            for lvl in [1;2;3] do
                use cmd = new Microsoft.Data.SqlClient.SqlCommand()
                cmd.Connection <- conn
                cmd.CommandText <- "WISECON_PSIKOTEST.dbo.SP_PaketSoalGroup"
                cmd.CommandType <- CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@NoGroup", req.NoGroup) |> ignore
                cmd.Parameters.AddWithValue("@Act", "DEL") |> ignore
                cmd.Parameters.AddWithValue("@lvl", lvl) |> ignore
                if lvl < 3 then
                    use rdr = cmd.ExecuteReader()
                    rdr.Close()
                else
                    let _ = cmd.ExecuteNonQuery()
                    ()
            this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Detail/Modify")>]
    member this.ModifyDtl ([<FromBody>] req: DtlRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandText <- "WISECON_PSIKOTEST.dbo.SP_PaketSoalGroupDtl"
        cmd.CommandType <- CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@SeqNo", (if req.SeqNo.HasValue then req.SeqNo.Value else 0L)) |> ignore
        cmd.Parameters.AddWithValue("@NoPaket", req.NoPaket) |> ignore
        cmd.Parameters.AddWithValue("@NoGroup", req.NoGroup) |> ignore
        cmd.Parameters.AddWithValue("@NoUrut", req.NoUrut) |> ignore
        cmd.Parameters.AddWithValue("@Judul", (if isNull req.Judul then "" else req.Judul)) |> ignore
        cmd.Parameters.AddWithValue("@Deskripsi", (if isNull req.Deskripsi then "" else req.Deskripsi)) |> ignore
        cmd.Parameters.AddWithValue("@IsDownload", (if req.IsDownload then 1 else 0)) |> ignore
        cmd.Parameters.AddWithValue("@bAktif", (if req.bAktif then 1 else 0)) |> ignore
        cmd.Parameters.AddWithValue("@MediaFileName", (if isNull req.MediaFileName then "NOMEDIA" else req.MediaFileName)) |> ignore
        cmd.Parameters.AddWithValue("@UrlMedia", (if isNull req.UrlMedia then "" else req.UrlMedia)) |> ignore
        cmd.Parameters.AddWithValue("@TipeMedia", (if isNull req.TipeMedia then "NOMEDIA" else req.TipeMedia)) |> ignore
        cmd.Parameters.AddWithValue("@User", (if isNull req.User then "" else req.User)) |> ignore
        let act = if req.SeqNo.HasValue && req.SeqNo.Value > 0L then "EDIT" else "ADD"
        cmd.Parameters.AddWithValue("@Act", act) |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let mutable seqNo = 0L
            if rdr.Read() then
                try seqNo <- rdr.GetInt64(0) with _ -> ()
            rdr.Close()
            this.Ok(box {| SeqNo = seqNo |})
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Detail/Delete")>]
    member this.DeleteDtl ([<FromBody>] req: DeleteDtlRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        conn.Open()
        try
            for lvl in [1;2;3] do
                use cmd = new Microsoft.Data.SqlClient.SqlCommand()
                cmd.Connection <- conn
                cmd.CommandText <- "WISECON_PSIKOTEST.dbo.SP_PaketSoalGroupDtl"
                cmd.CommandType <- CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@SeqNo", req.SeqNo) |> ignore
                cmd.Parameters.AddWithValue("@Act", "DEL") |> ignore
                cmd.Parameters.AddWithValue("@lvl", lvl) |> ignore
                if lvl < 3 then
                    use rdr = cmd.ExecuteReader()
                    rdr.Close()
                else
                    let _ = cmd.ExecuteNonQuery()
                    ()
            this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Detail/Copy")>]
    member this.CopyDtl ([<FromBody>] req: CopyDtlRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        conn.Open()
        try
            use chk = new Microsoft.Data.SqlClient.SqlCommand()
            chk.Connection <- conn
            chk.CommandText <- "SELECT ISNULL(bAktif,0) FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup WHERE NoGroup=@NoGroup"
            chk.CommandType <- CommandType.Text
            chk.Parameters.AddWithValue("@NoGroup", req.DestNoGroup) |> ignore
            let r = chk.ExecuteScalar()
            let aktif = if isNull r then 0 else (try Convert.ToInt32(r) with _ -> 0)
            if aktif = 1 then this.BadRequest(box {| error = "Group tujuan aktif" |})
            else
                use cmd = new Microsoft.Data.SqlClient.SqlCommand()
                cmd.Connection <- conn
                cmd.CommandText <- "WISECON_PSIKOTEST.dbo.SP_PaketSoalGroupDtl"
                cmd.CommandType <- CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@SeqNo", req.SeqNo) |> ignore
                cmd.Parameters.AddWithValue("@NoGroupTujuan", req.DestNoGroup) |> ignore
                cmd.Parameters.AddWithValue("@User", (if isNull req.User then "" else req.User)) |> ignore
                cmd.Parameters.AddWithValue("@Act", "COPY") |> ignore
                let _ = cmd.ExecuteNonQuery()
                this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Answer/Modify")>]
    member this.ModifyJawaban ([<FromBody>] req: JawabanRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandText <- "WISECON_PSIKOTEST.dbo.SP_PaketSoalGroupDtlJawaban"
        cmd.CommandType <- CommandType.StoredProcedure
        cmd.Parameters.AddWithValue("@NoPaket", req.NoPaket) |> ignore
        cmd.Parameters.AddWithValue("@NoGroup", req.NoGroup) |> ignore
        cmd.Parameters.AddWithValue("@NoUrut", req.NoUrut) |> ignore
        cmd.Parameters.AddWithValue("@NoJawaban", req.NoJawaban) |> ignore
        cmd.Parameters.AddWithValue("@PoinJawaban", req.PoinJawaban) |> ignore
        cmd.Parameters.AddWithValue("@Jawaban", (if isNull req.Jawaban then "" else req.Jawaban)) |> ignore
        cmd.Parameters.AddWithValue("@UrlMedia", (if isNull req.UrlMedia then "" else req.UrlMedia)) |> ignore
        cmd.Parameters.AddWithValue("@MediaFileName", (if isNull req.MediaFileName then "NOMEDIA" else req.MediaFileName)) |> ignore
        cmd.Parameters.AddWithValue("@TipeMedia", (if isNull req.TipeMedia then "NOMEDIA" else req.TipeMedia)) |> ignore
        cmd.Parameters.AddWithValue("@TextMedia", (if isNull req.TextMedia then "" else req.TextMedia)) |> ignore
        cmd.Parameters.AddWithValue("@User", (if isNull req.User then "" else req.User)) |> ignore
        cmd.Parameters.AddWithValue("@Act", "ADD") |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let mutable seqNo = 0L
            if rdr.Read() then
                try seqNo <- rdr.GetInt64(0) with _ -> ()
            rdr.Close()
            this.Ok(box {| SeqNo = seqNo |})
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Answer/Delete")>]
    member this.DeleteJawaban ([<FromBody>] req: DeleteJawabanRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        conn.Open()
        try
            for lvl in [1;2] do
                use cmd = new Microsoft.Data.SqlClient.SqlCommand()
                cmd.Connection <- conn
                cmd.CommandText <- "WISECON_PSIKOTEST.dbo.SP_PaketSoalGroupDtlJawaban"
                cmd.CommandType <- CommandType.StoredProcedure
                cmd.Parameters.AddWithValue("@SeqNo", req.SeqNo) |> ignore
                cmd.Parameters.AddWithValue("@Act", "DEL") |> ignore
                cmd.Parameters.AddWithValue("@lvl", lvl) |> ignore
                if lvl = 1 then
                    use rdr = cmd.ExecuteReader()
                    rdr.Close()
                else
                    let _ = cmd.ExecuteNonQuery()
                    ()
            this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Questions/Norma/List")>]
    member this.ListNorma ([<FromQuery>] noGroup: int64) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT SeqNo, Nama, BatasAtas, BatasBawah, UserInput, TimeInput FROM WISECON_PSIKOTEST.dbo.MS_NormaDtl WHERE NoGroup=@g ORDER BY SeqNo"
        cmd.Parameters.AddWithValue("@g", noGroup) |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let seqNo = rdr.GetInt64(0)
                let nama = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                let atas = if rdr.IsDBNull(2) then 0 else rdr.GetInt32(2)
                let bawah = if rdr.IsDBNull(3) then 0 else rdr.GetInt32(3)
                let userInput = if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                let timeInput = if rdr.IsDBNull(5) then DateTime.MinValue else rdr.GetDateTime(5)
                rows.Add(box {| seqNo = string seqNo; label = nama; atas = atas; bawah = bawah; userInput = userInput; time = timeInput |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Norma/Modify")>]
    member this.ModifyNorma ([<FromBody>] req: NormaRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        conn.Open()
        try
            use chk = new Microsoft.Data.SqlClient.SqlCommand()
            chk.Connection <- conn
            chk.CommandType <- CommandType.Text
            chk.CommandText <- "SELECT SUM(CASE WHEN name='UserEdit' THEN 1 ELSE 0 END) AS HasUserEdit, SUM(CASE WHEN name='TimeEdit' THEN 1 ELSE 0 END) AS HasTimeEdit FROM sys.columns WHERE object_id = OBJECT_ID('WISECON_PSIKOTEST.dbo.MS_NormaDtl')"
            use chkRdr = chk.ExecuteReader()
            let mutable hasUserEdit = false
            let mutable hasTimeEdit = false
            if chkRdr.Read() then
                let a = try chkRdr.GetInt32(0) with _ -> 0
                let b = try chkRdr.GetInt32(1) with _ -> 0
                hasUserEdit <- a > 0
                hasTimeEdit <- b > 0
            chkRdr.Close()

            let seqNo =
                if req.SeqNo.HasValue && req.SeqNo.Value > 0L then
                    use cmd = new Microsoft.Data.SqlClient.SqlCommand()
                    cmd.Connection <- conn
                    cmd.CommandType <- CommandType.Text
                    let editCols =
                        match hasUserEdit, hasTimeEdit with
                        | true, true -> ", UserEdit=@u, TimeEdit=GETDATE()"
                        | true, false -> ", UserEdit=@u"
                        | false, true -> ", TimeEdit=GETDATE()"
                        | false, false -> ""
                    cmd.CommandText <- "UPDATE WISECON_PSIKOTEST.dbo.MS_NormaDtl SET Nama=@n, BatasAtas=@a, BatasBawah=@b" + editCols + " WHERE SeqNo=@s; SELECT @s"
                    cmd.Parameters.AddWithValue("@n", (if isNull req.LabelNorma then "" else req.LabelNorma)) |> ignore
                    cmd.Parameters.AddWithValue("@a", (if req.BatasAtas.HasValue then req.BatasAtas.Value else 0)) |> ignore
                    cmd.Parameters.AddWithValue("@b", (if req.BatasBawah.HasValue then req.BatasBawah.Value else 0)) |> ignore
                    cmd.Parameters.AddWithValue("@u", (if isNull req.User then "" else req.User)) |> ignore
                    cmd.Parameters.AddWithValue("@s", req.SeqNo.Value) |> ignore
                    let v = cmd.ExecuteScalar()
                    if isNull v then 0L else Convert.ToInt64(v)
                else
                    use getMax = new Microsoft.Data.SqlClient.SqlCommand()
                    getMax.Connection <- conn
                    getMax.CommandType <- CommandType.Text
                    getMax.CommandText <- "SELECT ISNULL(MAX(SeqNo),0)+1 FROM WISECON_PSIKOTEST.dbo.MS_NormaDtl"
                    let nextObj = getMax.ExecuteScalar()
                    let nextSeq = if isNull nextObj then 1L else Convert.ToInt64(nextObj)
                    use ins = new Microsoft.Data.SqlClient.SqlCommand()
                    ins.Connection <- conn
                    ins.CommandType <- CommandType.Text
                    ins.CommandText <- "INSERT INTO WISECON_PSIKOTEST.dbo.MS_NormaDtl (SeqNo, Nama, NoGroup, BatasAtas, BatasBawah, UserInput, TimeInput) VALUES (@s, @n, @g, @a, @b, @u, GETDATE()); SELECT @s"
                    ins.Parameters.AddWithValue("@s", nextSeq) |> ignore
                    ins.Parameters.AddWithValue("@n", (if isNull req.LabelNorma then "" else req.LabelNorma)) |> ignore
                    ins.Parameters.AddWithValue("@g", req.NoGroup) |> ignore
                    ins.Parameters.AddWithValue("@a", (if req.BatasAtas.HasValue then req.BatasAtas.Value else 0)) |> ignore
                    ins.Parameters.AddWithValue("@b", (if req.BatasBawah.HasValue then req.BatasBawah.Value else 0)) |> ignore
                    ins.Parameters.AddWithValue("@u", (if isNull req.User then "" else req.User)) |> ignore
                    let v = ins.ExecuteScalar()
                    if isNull v then nextSeq else Convert.ToInt64(v)
            this.Ok(box {| SeqNo = seqNo |})
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Norma/Delete")>]
    member this.DeleteNorma ([<FromBody>] req: DeleteNormaRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "DELETE FROM WISECON_PSIKOTEST.dbo.MS_NormaDtl WHERE SeqNo=@s"
        cmd.Parameters.AddWithValue("@s", req.SeqNo) |> ignore
        conn.Open()
        try
            let _ = cmd.ExecuteNonQuery()
            this.Ok()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Questions/Paket/List")>]
    member this.ListPaket ([<FromQuery>] fromDate: string, [<FromQuery>] toDate: string, [<FromQuery>] status: string) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        let baseSelect = "SELECT ps.NoPaket, ps.NamaPaket, ps.ToleransiWaktu, ps.bAktif, ps.UserInput, ps.TimeInput, ps.UserEdit, ps.TimeEdit, {POS} FROM WISECON_PSIKOTEST.dbo.MS_PaketSoal ps WHERE ( @st IS NULL OR @st = '' OR ps.bAktif = CASE WHEN @st='AKTIF' THEN 1 ELSE 0 END ) AND ( (@f IS NULL OR @f = '') OR (@t IS NULL OR @t = '') OR (ps.TimeInput BETWEEN CONVERT(datetime,@f+'T00:00:00') AND CONVERT(datetime,@t+'T23:59:59')) ) ORDER BY ps.NoPaket DESC"
        let posExpr =
            let conn2 = db :?> Microsoft.Data.SqlClient.SqlConnection
            use chk = new Microsoft.Data.SqlClient.SqlCommand()
            chk.Connection <- conn2
            chk.CommandType <- CommandType.Text
            chk.CommandText <- "SELECT CASE WHEN EXISTS (SELECT 1 FROM sys.columns WHERE object_id = OBJECT_ID('WISECON_PSIKOTEST.dbo.MS_PaketSoal') AND name = 'Position') THEN 1 ELSE 0 END"
            conn2.Open()
            let vObj = chk.ExecuteScalar()
            conn2.Close()
            let hasPos = try Convert.ToInt32(vObj) = 1 with _ -> false
            if hasPos then "ISNULL(CAST(ps.Position AS NVARCHAR(200)), '') AS Position" else "'' AS Position"
        cmd.CommandText <- baseSelect.Replace("{POS}", posExpr)
        cmd.Parameters.AddWithValue("@f", (if String.IsNullOrWhiteSpace(fromDate) then box DBNull.Value else box fromDate)) |> ignore
        cmd.Parameters.AddWithValue("@t", (if String.IsNullOrWhiteSpace(toDate) then box DBNull.Value else box toDate)) |> ignore
        cmd.Parameters.AddWithValue("@st", (if String.IsNullOrWhiteSpace(status) then box DBNull.Value else box status)) |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let noPaket = rdr.GetInt32(0)
                let namaPaket = rdr.GetString(1)
                let toleransi = rdr.GetInt32(2)
                let bAktif = try rdr.GetBoolean(3) with _ -> false
                let userInput = if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                let timeInput = if rdr.IsDBNull(5) then DateTime.MinValue else rdr.GetDateTime(5)
                let userEdit = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                let timeEdit = if rdr.IsDBNull(7) then DateTime.MinValue else rdr.GetDateTime(7)
                let posisi = if rdr.IsDBNull(8) then "" else rdr.GetString(8)
                rows.Add(box {| noPaket = string noPaket; namaPaket = namaPaket; posisi = posisi; toleransi = toleransi; status = (if bAktif then "AKTIF" else "NONAKTIF"); userInput = userInput; time = timeInput; userEdit = userEdit; timeEdit = timeEdit |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Questions/Group/List")>]
    member this.ListGroup ([<FromQuery>] noPaket: int64) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <-
            "SELECT NoGroup, NoPaket, NamaGroup, MinimumJmlSoal, NilaiStandar, WaktuPengerjaan, bRandom, bAktif, NoPetunjuk, UserInput, TimeInput, UserEdit, TimeEdit, IsPrioritas FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup WHERE NoPaket=@p ORDER BY NoGroup"
        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let noGroup = rdr.GetInt32(0)
                let namaGroup = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                let minSoal = if rdr.IsDBNull(3) then 0 else rdr.GetInt32(3)
                let nilaiStandar = if rdr.IsDBNull(4) then 0 else rdr.GetInt32(4)
                let waktu = if rdr.IsDBNull(5) then 0 else rdr.GetInt32(5)
                let bRandom = try rdr.GetBoolean(6) with _ -> false
                let bAktif = try rdr.GetBoolean(7) with _ -> false
                let petunjuk = if rdr.IsDBNull(8) then 0 else rdr.GetInt32(8)
                let userInput = if rdr.IsDBNull(9) then "" else rdr.GetString(9)
                let timeInput = if rdr.IsDBNull(10) then DateTime.MinValue else rdr.GetDateTime(10)
                let userEdit = if rdr.IsDBNull(11) then "" else rdr.GetString(11)
                let timeEdit = if rdr.IsDBNull(12) then DateTime.MinValue else rdr.GetDateTime(12)
                let isPrioritas = if rdr.FieldCount > 13 then (try rdr.GetBoolean(13) with _ -> false) else false
                rows.Add(box {| noGroup = string noGroup; namaGroup = namaGroup; minSoal = minSoal; nilaiStandar = nilaiStandar; waktu = waktu; random = (if bRandom then "Ya" else "Tidak"); status = (if bAktif then "AKTIF" else "NONAKTIF"); petunjuk = petunjuk; userInput = userInput; time = timeInput; userEdit = userEdit; timeEdit = timeEdit; isPrioritas = isPrioritas |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Questions/Detail/List")>]
    member this.ListDtl ([<FromQuery>] noPaket: int64, [<FromQuery>] noGroup: int64) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <-
            "SELECT SeqNo, NoUrut, Judul, Deskripsi, IsDownload, bAktif, UrlMedia, TipeMedia, UserInput, TimeInput, UserEdit, TimeEdit FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtl WHERE NoPaket=@p AND NoGroup=@g ORDER BY NoUrut"
        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
        cmd.Parameters.AddWithValue("@g", noGroup) |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let seqNo = rdr.GetInt64(0)
                let noUrut = rdr.GetInt32(1)
                let judul = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                let deskripsi = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                let isDownload = try rdr.GetBoolean(4) with _ -> false
                let aktif = try rdr.GetBoolean(5) with _ -> false
                let urlMedia = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                let tipeMedia = if rdr.IsDBNull(7) then "NOMEDIA" else rdr.GetString(7)
                let userInput = if rdr.IsDBNull(8) then "" else rdr.GetString(8)
                let timeInput = if rdr.IsDBNull(9) then DateTime.MinValue else rdr.GetDateTime(9)
                let userEdit = if rdr.IsDBNull(10) then "" else rdr.GetString(10)
                let timeEdit = if rdr.IsDBNull(11) then DateTime.MinValue else rdr.GetDateTime(11)
                rows.Add(box {| seqNo = string seqNo; noUrut = int noUrut; judul = judul; deskripsi = deskripsi; isDownload = (if isDownload then "Ya" else "Tidak"); aktif = (if aktif then "Ya" else "Tidak"); urlMedia = urlMedia; tipeMedia = tipeMedia; userInput = userInput; time = timeInput; userEdit = userEdit; timeEdit = timeEdit; noSoal = string seqNo |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Questions/Answer/List")>]
    member this.ListJawaban ([<FromQuery>] noPaket: int64, [<FromQuery>] noGroup: int64, [<FromQuery>] noUrut: int64) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <-
            "SELECT SeqNo, NoJawaban, Jawaban, NoJawabanBenar, TipeMedia, UrlMedia, UserInput, TimeInput, UserEdit, TimeEdit FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtlJawaban WHERE NoPaket=@p AND NoGroup=@g AND NoUrut=@u ORDER BY NoJawaban"
        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
        cmd.Parameters.AddWithValue("@g", noGroup) |> ignore
        cmd.Parameters.AddWithValue("@u", noUrut) |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let seqNo = rdr.GetInt64(0)
                let noJawaban = if rdr.IsDBNull(1) then 0 else int (rdr.GetByte(1))
                let jawaban = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                let poin = if rdr.IsDBNull(3) then 0 else int (rdr.GetInt16(3))
                let tipeMedia = if rdr.IsDBNull(4) then "NOMEDIA" else rdr.GetString(4)
                let urlMedia = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                let userInput = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                let timeInput = if rdr.IsDBNull(7) then DateTime.MinValue else rdr.GetDateTime(7)
                let userEdit = if rdr.IsDBNull(8) then "" else rdr.GetString(8)
                let timeEdit = if rdr.IsDBNull(9) then DateTime.MinValue else rdr.GetDateTime(9)
                rows.Add(box {| seqNo = string seqNo; noJawaban = string noJawaban; jawaban = jawaban; poin = poin; tipeMedia = tipeMedia; urlMedia = urlMedia; userInput = userInput; time = timeInput; userEdit = userEdit; timeEdit = timeEdit |})
            this.Ok(rows)
        finally
            conn.Close()
    [<Authorize>]
    [<HttpPost>]
    [<Route("Questions/Detail/UploadMedia")>]
    member this.UploadMedia([<FromForm>] file: IFormFile) : IActionResult =
        if isNull file || file.Length <= 0L then
            this.BadRequest(box {| error = "File kosong" |})
        else
            let envObj = this.HttpContext.RequestServices.GetService(typeof<Microsoft.AspNetCore.Hosting.IWebHostEnvironment>)
            let env = envObj :?> Microsoft.AspNetCore.Hosting.IWebHostEnvironment
            let uploads = System.IO.Path.Combine(env.WebRootPath, "uploads")
            if not (System.IO.Directory.Exists(uploads)) then System.IO.Directory.CreateDirectory(uploads) |> ignore
            let name = sprintf "%s_%s%s" (System.DateTime.UtcNow.ToString("yyyyMMddHHmmss")) (System.Guid.NewGuid().ToString("N").Substring(0,8)) (System.IO.Path.GetExtension(file.FileName))
            let path = System.IO.Path.Combine(uploads, name)
            use fs = new System.IO.FileStream(path, System.IO.FileMode.Create)
            file.CopyTo(fs)
            let url = "/uploads/" + name
            this.Ok(box {| url = url; fileName = name |})
