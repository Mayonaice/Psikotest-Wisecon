namespace PsikotestWisesa.Controllers

open System
open System.Data
open System.IO
open System.Text.Json
open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authorization
open Microsoft.AspNetCore.Http
open ClosedXML.Excel

[<CLIMutable>]
type ApproveItem = { kodeBiodata: string }

[<CLIMutable>]
type ApproveRequest = { items: ApproveItem array }

[<CLIMutable>]
type InterviewItem = { msSeqNo: int; question: string; answer: string; remark: string }

[<CLIMutable>]
type PersonalityItem = { msSeqNo: int; aspek: string; uraian: string; nilai: string }

[<CLIMutable>]
type InterviewSavePayload = { personal: InterviewItem array; experience: InterviewItem array; personality: PersonalityItem array }

[<ApiController>]
type ScreeningController (db: IDbConnection) =
    inherit Controller()

    member private this.getQueryValue(key: string) =
        if this.Request.Query.ContainsKey(key) then this.Request.Query.[key].ToString() else ""

    member private this.applyFilters(cmd: Microsoft.Data.SqlClient.SqlCommand, ?alias: string) =
        let prefix =
            match alias with
            | Some a when not (String.IsNullOrWhiteSpace(a)) -> a + "."
            | _ -> ""
        let mutable whereClause = " WHERE 1=1 "
        let fromDateStr = this.getQueryValue("fromDate")
        if not (String.IsNullOrWhiteSpace(fromDateStr)) then
            match DateTime.TryParse(fromDateStr) with
            | true, v ->
                whereClause <- whereClause + " AND " + prefix + "TimeInput >= @fromDate "
                cmd.Parameters.AddWithValue("@fromDate", v.Date) |> ignore
            | _ -> ()
        let toDateStr = this.getQueryValue("toDate")
        if not (String.IsNullOrWhiteSpace(toDateStr)) then
            match DateTime.TryParse(toDateStr) with
            | true, v ->
                whereClause <- whereClause + " AND " + prefix + "TimeInput <= @toDate "
                cmd.Parameters.AddWithValue("@toDate", v.Date.AddDays(1.0).AddSeconds(-1.0)) |> ignore
            | _ -> ()
        let statusSeleksi = this.getQueryValue("statusSeleksi")
        if not (String.IsNullOrWhiteSpace(statusSeleksi)) then
            whereClause <- whereClause + " AND " + prefix + "StatusScreening = @statusSeleksi "
            cmd.Parameters.AddWithValue("@statusSeleksi", statusSeleksi) |> ignore
        let statusVerifikasi = this.getQueryValue("statusVerifikasi")
        if not (String.IsNullOrWhiteSpace(statusVerifikasi)) then
            whereClause <- whereClause + " AND " + prefix + "StatusVerif = @statusVerifikasi "
            cmd.Parameters.AddWithValue("@statusVerifikasi", statusVerifikasi) |> ignore
        whereClause

    member private this.getUserName() =
        let name = if isNull this.User || isNull this.User.Identity then "" else this.User.Identity.Name
        if isNull name then "" else name

    member private this.saveUpload(file: IFormFile, folder: string) =
        if isNull file || file.Length <= 0L then
            ""
        else
            if not (Directory.Exists(folder)) then Directory.CreateDirectory(folder) |> ignore
            let ext = Path.GetExtension(file.FileName)
            let name = sprintf "%s_%s%s" (DateTime.UtcNow.ToString("yyyyMMddHHmmss")) (Guid.NewGuid().ToString("N").Substring(0,8)) ext
            let path = Path.Combine(folder, name)
            use fs = new FileStream(path, FileMode.Create)
            file.CopyTo(fs)
            name

    [<Authorize>]
    [<HttpGet>]
    [<Route("Screening/Top/List")>]
    member this.TopList () : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        let whereClause = this.applyFilters(cmd)
        cmd.CommandText <-
            "SELECT MIN(JobCode) AS JobCode, JobPosition, " +
            "SUM(CASE WHEN StatusScreening IS NULL OR LTRIM(RTRIM(StatusScreening))='' OR StatusScreening='Belum Proses' THEN 1 ELSE 0 END) AS BelumProses, " +
            "SUM(CASE WHEN StatusScreening='Interview Proses' THEN 1 ELSE 0 END) AS InterviewProses, " +
            "SUM(CASE WHEN StatusScreening='Interview Tidak Datang' THEN 1 ELSE 0 END) AS InterviewTidakDatang, " +
            "SUM(CASE WHEN StatusScreening='Psikotest Proses' THEN 1 ELSE 0 END) AS PsikotestProses, " +
            "SUM(CASE WHEN StatusScreening='Psikotest Tidak Datang' THEN 1 ELSE 0 END) AS PsikotestTidakDatang, " +
            "SUM(CASE WHEN StatusVerif='Lulus' THEN 1 ELSE 0 END) AS StatusLulus, " +
            "SUM(CASE WHEN StatusVerif='Tidak Lulus' THEN 1 ELSE 0 END) AS StatusTidakLulus, " +
            "MAX(TimeInput) AS TimeInput " +
            "FROM WISECON_PSIKOTEST.dbo.VW_REC_Screening " +
            whereClause +
            "GROUP BY JobPosition " +
            "ORDER BY JobPosition"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let code = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                let name = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                let belum = if rdr.IsDBNull(2) then 0 else rdr.GetInt32(2)
                let intProses = if rdr.IsDBNull(3) then 0 else rdr.GetInt32(3)
                let intTidak = if rdr.IsDBNull(4) then 0 else rdr.GetInt32(4)
                let psiProses = if rdr.IsDBNull(5) then 0 else rdr.GetInt32(5)
                let psiTidak = if rdr.IsDBNull(6) then 0 else rdr.GetInt32(6)
                let lulus = if rdr.IsDBNull(7) then 0 else rdr.GetInt32(7)
                let tdk = if rdr.IsDBNull(8) then 0 else rdr.GetInt32(8)
                let timeInput = if rdr.IsDBNull(9) then DateTime.MinValue else rdr.GetDateTime(9)
                rows.Add(box {| code = code; name = name; key = name; belumProses = belum; interviewProses = intProses; interviewTidakDatang = intTidak; psikotestProses = psiProses; psikotestTidakDatang = psiTidak; statusLulus = lulus; statusTidakLulus = tdk; timeInput = timeInput |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Screening/Top/Export")>]
    member this.TopExport () : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        let whereClause = this.applyFilters(cmd, "v")
        cmd.CommandText <-
            "WITH Ranked AS (" +
            "SELECT v.SeqNo, v.JobCode, v.JobPosition, v.BatchNo, v.Name, v.StatusScreening, v.StatusVerif, v.ResultInterviewTes, v.StatusPsikotes, v.Sex, v.UserEmail, v.PhoneNo, v.PendidikanTerakhir, v.TimeInput, " +
            "ROW_NUMBER() OVER (PARTITION BY ISNULL(p.NoPeserta, CONCAT('SEQ_', v.SeqNo)) ORDER BY v.TimeInput DESC, v.SeqNo DESC) AS rn " +
            "FROM WISECON_PSIKOTEST.dbo.VW_REC_Screening v " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.MS_Peserta p ON p.ID = v.SeqNo " +
            whereClause +
            ") " +
            "SELECT JobPosition, BatchNo, Name, StatusScreening, StatusVerif, ResultInterviewTes, StatusPsikotes, Sex, UserEmail, PhoneNo, PendidikanTerakhir, TimeInput " +
            "FROM Ranked WHERE rn=1 ORDER BY Name"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            use wb = new XLWorkbook()
            let ws = wb.AddWorksheet("Screening")
            ws.Cell(1,1).Value <- "JobPosition"
            ws.Cell(1,2).Value <- "Batch"
            ws.Cell(1,3).Value <- "Name"
            ws.Cell(1,4).Value <- "StatusScreening"
            ws.Cell(1,5).Value <- "StatusVerif"
            ws.Cell(1,6).Value <- "ResultInterview"
            ws.Cell(1,7).Value <- "StatusPsikotes"
            ws.Cell(1,8).Value <- "Sex"
            ws.Cell(1,9).Value <- "Email"
            ws.Cell(1,10).Value <- "Phone"
            ws.Cell(1,11).Value <- "PendidikanTerakhir"
            ws.Cell(1,12).Value <- "TimeInput"
            let mutable row = 2
            while rdr.Read() do
                ws.Cell(row,1).Value <- if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                ws.Cell(row,2).Value <- if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                ws.Cell(row,3).Value <- if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                ws.Cell(row,4).Value <- if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                ws.Cell(row,5).Value <- if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                ws.Cell(row,6).Value <- if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                ws.Cell(row,7).Value <- if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                ws.Cell(row,8).Value <- if rdr.IsDBNull(7) then "" else rdr.GetString(7)
                ws.Cell(row,9).Value <- if rdr.IsDBNull(8) then "" else rdr.GetString(8)
                ws.Cell(row,10).Value <- if rdr.IsDBNull(9) then "" else rdr.GetString(9)
                ws.Cell(row,11).Value <- if rdr.IsDBNull(10) then "" else rdr.GetString(10)
                let t = if rdr.IsDBNull(11) then DateTime.MinValue else rdr.GetDateTime(11)
                ws.Cell(row,12).Value <- (if t = DateTime.MinValue then "" else t.ToString("yyyy-MM-dd HH:mm"))
                row <- row + 1
            ws.Columns().AdjustToContents() |> ignore
            use ms = new MemoryStream()
            wb.SaveAs(ms)
            this.File(ms.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Screening.xlsx")
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Screening/Biodata/List")>]
    member this.BiodataList () : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        let whereClause = this.applyFilters(cmd, "v")
        let kodeLowongan = this.getQueryValue("kodeLowongan")
        if not (String.IsNullOrWhiteSpace(kodeLowongan)) then
            cmd.Parameters.AddWithValue("@kodeLowongan", kodeLowongan) |> ignore
        cmd.CommandText <-
            "WITH Ranked AS (" +
            "SELECT v.SeqNo, v.JobCode, v.JobPosition, v.BatchNo, v.Name, v.StatusScreening, v.StatusVerif, v.ResultInterviewTes, v.StatusPsikotes, v.Sex, v.UserEmail, v.PhoneNo, v.PendidikanTerakhir, v.CVFileName, v.TimeInput, " +
            "ROW_NUMBER() OVER (PARTITION BY ISNULL(p.NoPeserta, CONCAT('SEQ_', v.SeqNo)) ORDER BY v.TimeInput DESC, v.SeqNo DESC) AS rn " +
            "FROM WISECON_PSIKOTEST.dbo.VW_REC_Screening v " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.MS_Peserta p ON p.ID = v.SeqNo " +
            whereClause +
            (if String.IsNullOrWhiteSpace(kodeLowongan) then "" else " AND v.JobPosition = @kodeLowongan ") +
            ") " +
            "SELECT SeqNo, JobCode, JobPosition, BatchNo, Name, StatusScreening, StatusVerif, ResultInterviewTes, StatusPsikotes, Sex, UserEmail, PhoneNo, PendidikanTerakhir, CVFileName, TimeInput " +
            "FROM Ranked WHERE rn=1 ORDER BY Name"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let seqNo = if rdr.IsDBNull(0) then 0 else rdr.GetInt32(0)
                let jobCode = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                let jobPos = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                let batch = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                let name = if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                let statusScreening = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                let statusVerif = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                let hasilInterview = if rdr.IsDBNull(7) then "" else rdr.GetString(7)
                let hasilPsikotest = if rdr.IsDBNull(8) then "" else rdr.GetString(8)
                let sex = if rdr.IsDBNull(9) then "" else rdr.GetString(9)
                let email = if rdr.IsDBNull(10) then "" else rdr.GetString(10)
                let phone = if rdr.IsDBNull(11) then "" else rdr.GetString(11)
                let pendidikan = if rdr.IsDBNull(12) then "" else rdr.GetString(12)
                let cv = if rdr.IsDBNull(13) then "" else rdr.GetString(13)
                let timeInput = if rdr.IsDBNull(14) then DateTime.MinValue else rdr.GetDateTime(14)
                rows.Add(box {| kodeLowongan = jobCode; kodeBiodata = string seqNo; posLowongan = jobPos; batch = batch; name = name; statusScreening = statusScreening; statusVerifikasi = statusVerif; hasilInterview = hasilInterview; hasilPsikotest = hasilPsikotest; sex = sex; email = email; phone = phone; pendidikanTerakhir = pendidikan; cvFileName = cv; timeInput = timeInput |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Screening/Biodata/DownloadNB")>]
    member this.DownloadBiodata([<FromQuery>] kodeBiodata: int) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT CVFileName FROM WISECON_PSIKOTEST.dbo.MS_Peserta WHERE ID=@id", conn)
        cmd.Parameters.AddWithValue("@id", kodeBiodata) |> ignore
        conn.Open()
        try
            let cvObj = cmd.ExecuteScalar()
            let cv = if isNull cvObj then "" else cvObj.ToString()
            if String.IsNullOrWhiteSpace(cv) then this.NotFound() :> IActionResult
            else
                let basePath = @"C:\dct_docs\WISECON_PSIKOTEST\Resume"
                let fileName = Path.GetFileName(cv)
                let filePath = Path.Combine(basePath, fileName)
                if not (System.IO.File.Exists(filePath)) then this.NotFound() :> IActionResult
                else this.PhysicalFile(filePath, "application/octet-stream", fileName) :> IActionResult
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Screening/Biodata/Interview/Download")>]
    member this.DownloadInterview([<FromQuery>] kodeBiodata: int) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT InterviewResultFileName FROM WISECON_PSIKOTEST.dbo.REC_ScreeningStatus WHERE JobVacancySubmittedSeqNo=@id", conn)
        cmd.Parameters.AddWithValue("@id", kodeBiodata) |> ignore
        conn.Open()
        try
            let nameObj = cmd.ExecuteScalar()
            let name = if isNull nameObj then "" else nameObj.ToString()
            if String.IsNullOrWhiteSpace(name) then this.NotFound() :> IActionResult
            else
                let basePath = @"C:\dct_docs\WISECON_PSIKOTEST\Screening\Interview"
                let fileName = Path.GetFileName(name)
                let filePath = Path.Combine(basePath, fileName)
                if not (System.IO.File.Exists(filePath)) then this.NotFound() :> IActionResult
                else this.PhysicalFile(filePath, "application/octet-stream", fileName) :> IActionResult
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Screening/Biodata/Psikotest/Download")>]
    member this.DownloadPsikotest([<FromQuery>] kodeBiodata: int) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT PsikotesResultFileName FROM WISECON_PSIKOTEST.dbo.REC_ScreeningStatus WHERE JobVacancySubmittedSeqNo=@id", conn)
        cmd.Parameters.AddWithValue("@id", kodeBiodata) |> ignore
        conn.Open()
        try
            let nameObj = cmd.ExecuteScalar()
            let name = if isNull nameObj then "" else nameObj.ToString()
            if String.IsNullOrWhiteSpace(name) then this.NotFound() :> IActionResult
            else
                let basePath = @"C:\dct_docs\WISECON_PSIKOTEST\Screening\Psikotest"
                let fileName = Path.GetFileName(name)
                let filePath = Path.Combine(basePath, fileName)
                if not (System.IO.File.Exists(filePath)) then this.NotFound() :> IActionResult
                else this.PhysicalFile(filePath, "application/octet-stream", fileName) :> IActionResult
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Screening/Biodata/Paper")>]
    member this.Paper([<FromQuery>] kodeBiodata: int, [<FromQuery>] noPeserta: string) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        let options = JsonSerializerOptions(PropertyNamingPolicy = JsonNamingPolicy.CamelCase)
        options.PropertyNameCaseInsensitive <- true
        let noPesertaInput = if isNull noPeserta then "" else noPeserta.Trim()
        let mutable biodataId = kodeBiodata
        conn.Open()
        try
            if not (String.IsNullOrWhiteSpace(noPesertaInput)) then
                use cmdId = new Microsoft.Data.SqlClient.SqlCommand("SELECT TOP 1 ID FROM WISECON_PSIKOTEST.dbo.MS_Peserta WHERE NoPeserta=@np", conn)
                cmdId.Parameters.AddWithValue("@np", noPesertaInput) |> ignore
                let idObj = cmdId.ExecuteScalar()
                if not (isNull idObj) then
                    biodataId <- Convert.ToInt32(idObj)
            let mutable noPesertaBio = noPesertaInput
            let mutable noKtpBio = ""
            use cmdPeserta = new Microsoft.Data.SqlClient.SqlCommand("SELECT TOP 1 NoPeserta, NoKTP FROM WISECON_PSIKOTEST.dbo.MS_Peserta WHERE ID=@id", conn)
            cmdPeserta.Parameters.AddWithValue("@id", biodataId) |> ignore
            use rdrPeserta = cmdPeserta.ExecuteReader()
            if rdrPeserta.Read() then
                noPesertaBio <- if rdrPeserta.IsDBNull(0) then noPesertaBio else rdrPeserta.GetString(0)
                noKtpBio <- if rdrPeserta.IsDBNull(1) then "" else rdrPeserta.GetString(1)
            rdrPeserta.Close()

            use cmdBio = new Microsoft.Data.SqlClient.SqlCommand("SELECT SeqNo, Name, UserEmail, JobPosition, PhoneNo, PendidikanTerakhir, StatusScreening, StatusVerif, ResultInterviewTes, StatusPsikotes FROM WISECON_PSIKOTEST.dbo.VW_REC_Screening WHERE SeqNo=@id", conn)
            cmdBio.Parameters.AddWithValue("@id", biodataId) |> ignore
            use rdrBio = cmdBio.ExecuteReader()
            let mutable biodataObj = box {| seqNo = biodataId; noPeserta = noPesertaBio; noKtp = noKtpBio; name = ""; email = ""; job = ""; phone = ""; pendidikan = ""; statusScreening = ""; statusVerif = ""; hasilInterview = ""; hasilPsikotest = "" |}
            let mutable bioEmail = ""
            if rdrBio.Read() then
                let name = if rdrBio.IsDBNull(1) then "" else rdrBio.GetString(1)
                let email = if rdrBio.IsDBNull(2) then "" else rdrBio.GetString(2)
                let job = if rdrBio.IsDBNull(3) then "" else rdrBio.GetString(3)
                let phone = if rdrBio.IsDBNull(4) then "" else rdrBio.GetString(4)
                let pendidikan = if rdrBio.IsDBNull(5) then "" else rdrBio.GetString(5)
                let statusScreening = if rdrBio.IsDBNull(6) then "" else rdrBio.GetString(6)
                let statusVerif = if rdrBio.IsDBNull(7) then "" else rdrBio.GetString(7)
                let hasilInterview = if rdrBio.IsDBNull(8) then "" else rdrBio.GetString(8)
                let hasilPsikotest = if rdrBio.IsDBNull(9) then "" else rdrBio.GetString(9)
                bioEmail <- email
                biodataObj <- box {| seqNo = biodataId; noPeserta = noPesertaBio; noKtp = noKtpBio; name = name; email = email; job = job; phone = phone; pendidikan = pendidikan; statusScreening = statusScreening; statusVerif = statusVerif; hasilInterview = hasilInterview; hasilPsikotest = hasilPsikotest |}
            rdrBio.Close()

            let loadList (sql: string) (mapFn: IDataRecord -> obj) =
                use cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, conn)
                use rdr = cmd.ExecuteReader()
                let rows = ResizeArray<obj>()
                while rdr.Read() do
                    rows.Add(mapFn rdr)
                rdr.Close()
                rows

            let loadListById (sql: string) (mapFn: IDataRecord -> obj) =
                use cmd = new Microsoft.Data.SqlClient.SqlCommand(sql, conn)
                cmd.Parameters.AddWithValue("@id", biodataId) |> ignore
                use rdr = cmd.ExecuteReader()
                let rows = ResizeArray<obj>()
                while rdr.Read() do
                    rows.Add(mapFn rdr)
                rdr.Close()
                rows

            let personal = loadList "SELECT SeqNo, Question FROM WISECON_PSIKOTEST.dbo.REC_MsPersonalInterview ORDER BY SeqNo" (fun r ->
                let ms = if r.IsDBNull(0) then 0 else Convert.ToInt32(r.GetValue(0))
                let q = if r.IsDBNull(1) then "" else r.GetString(1)
                box {| msSeqNo = ms; question = q |})
            let experience = loadList "SELECT SeqNo, Question FROM WISECON_PSIKOTEST.dbo.REC_MsExperienceInterview ORDER BY SeqNo" (fun r ->
                let ms = if r.IsDBNull(0) then 0 else Convert.ToInt32(r.GetValue(0))
                let q = if r.IsDBNull(1) then "" else r.GetString(1)
                box {| msSeqNo = ms; question = q |})
            let personality = loadList "SELECT SeqNo, Aspek, Uraian FROM WISECON_PSIKOTEST.dbo.REC_MsPersonalityTest ORDER BY SeqNo" (fun r ->
                let ms = if r.IsDBNull(0) then 0 else Convert.ToInt32(r.GetValue(0))
                let a = if r.IsDBNull(1) then "" else r.GetString(1)
                let u = if r.IsDBNull(2) then "" else r.GetString(2)
                box {| msSeqNo = ms; aspek = a; uraian = u |})

            let personalRes = loadListById "SELECT MsSeqNo, Answer, Remark FROM WISECON_PSIKOTEST.dbo.REC_ResultPersonalInterview WHERE JobVacancySubmittedSeqNo=@id ORDER BY MsSeqNo" (fun r ->
                let ms = r.GetInt32(0)
                let a = if r.IsDBNull(1) then "" else r.GetString(1)
                let rmk = if r.IsDBNull(2) then "" else r.GetString(2)
                box {| msSeqNo = ms; answer = a; remark = rmk |})
            let experienceRes = loadListById "SELECT MsSeqNo, Answer, Remark FROM WISECON_PSIKOTEST.dbo.REC_ResultExperienceInterview WHERE JobVacancySubmittedSeqNo=@id ORDER BY MsSeqNo" (fun r ->
                let ms = r.GetInt32(0)
                let a = if r.IsDBNull(1) then "" else r.GetString(1)
                let rmk = if r.IsDBNull(2) then "" else r.GetString(2)
                box {| msSeqNo = ms; answer = a; remark = rmk |})
            let personalityRes = loadListById "SELECT MsSeqNo, Nilai FROM WISECON_PSIKOTEST.dbo.REC_ResultPersonalityTest WHERE JobVacancySubmittedSeqNo=@id ORDER BY MsSeqNo" (fun r ->
                let ms = r.GetInt32(0)
                let n = if r.IsDBNull(1) then "" else r.GetString(1)
                box {| msSeqNo = ms; nilai = n |})

            let statusCmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT ResultInterviewTes, InterviewResultRemarks, InterviewResultFileName, StatusPsikotes, PsikotesResultFileName, InterviewStrengthWeakness, InterviewLikeDislike FROM WISECON_PSIKOTEST.dbo.REC_ScreeningStatus WHERE JobVacancySubmittedSeqNo=@id", conn)
            statusCmd.Parameters.AddWithValue("@id", biodataId) |> ignore
            use statusR = statusCmd.ExecuteReader()
            let mutable statusObj = box {| hasilInterview = ""; remarks = ""; interviewFile = ""; hasilPsikotest = ""; psikotestFile = ""; strengthWeakness = ""; likeDislike = "" |}
            if statusR.Read() then
                let hasilInterview = if statusR.IsDBNull(0) then "" else statusR.GetString(0)
                let remarks = if statusR.IsDBNull(1) then "" else statusR.GetString(1)
                let interviewFile = if statusR.IsDBNull(2) then "" else statusR.GetString(2)
                let hasilPsikotest = if statusR.IsDBNull(3) then "" else statusR.GetString(3)
                let psikotestFile = if statusR.IsDBNull(4) then "" else statusR.GetString(4)
                let strengthWeakness = if statusR.IsDBNull(5) then "" else statusR.GetString(5)
                let likeDislike = if statusR.IsDBNull(6) then "" else statusR.GetString(6)
                statusObj <- box {| hasilInterview = hasilInterview; remarks = remarks; interviewFile = interviewFile; hasilPsikotest = hasilPsikotest; psikotestFile = psikotestFile; strengthWeakness = strengthWeakness; likeDislike = likeDislike |}
            statusR.Close()

            let mutable userId = ""
            let mutable noPaket = 0
            let mutable noPesertaDb = ""
            use cmdUser = new Microsoft.Data.SqlClient.SqlCommand("SELECT TOP 1 D.UserId, D.NoPaket, P.NoPeserta FROM WISECON_PSIKOTEST.dbo.MS_Peserta P LEFT JOIN WISECON_PSIKOTEST.dbo.MS_PesertaDtl D ON D.NoPeserta = P.NoPeserta WHERE P.ID=@id ORDER BY D.TimeInput DESC", conn)
            cmdUser.Parameters.AddWithValue("@id", biodataId) |> ignore
            use rdrUser = cmdUser.ExecuteReader()
            if rdrUser.Read() then
                userId <- if rdrUser.IsDBNull(0) then "" else rdrUser.GetString(0)
                noPaket <- if rdrUser.IsDBNull(1) then 0 else Convert.ToInt32(rdrUser.GetValue(1))
                noPesertaDb <- if rdrUser.IsDBNull(2) then "" else rdrUser.GetString(2)
            rdrUser.Close()
            if (String.IsNullOrWhiteSpace(userId) || noPaket <= 0) && not (String.IsNullOrWhiteSpace(bioEmail)) then
                use cmdUserAlt = new Microsoft.Data.SqlClient.SqlCommand("SELECT TOP 1 D.UserId, D.NoPaket, P.NoPeserta FROM WISECON_PSIKOTEST.dbo.MS_Peserta P LEFT JOIN WISECON_PSIKOTEST.dbo.MS_PesertaDtl D ON D.NoPeserta = P.NoPeserta WHERE P.Email=@email ORDER BY D.TimeInput DESC", conn)
                cmdUserAlt.Parameters.AddWithValue("@email", bioEmail) |> ignore
                use rdrUserAlt = cmdUserAlt.ExecuteReader()
                if rdrUserAlt.Read() then
                    userId <- if rdrUserAlt.IsDBNull(0) then "" else rdrUserAlt.GetString(0)
                    noPaket <- if rdrUserAlt.IsDBNull(1) then 0 else Convert.ToInt32(rdrUserAlt.GetValue(1))
                    noPesertaDb <- if rdrUserAlt.IsDBNull(2) then "" else rdrUserAlt.GetString(2)
                rdrUserAlt.Close()

            let altUserIds = ResizeArray<string * int>()
            if not (String.IsNullOrWhiteSpace(noPesertaDb)) then
                use cmdUserList = new Microsoft.Data.SqlClient.SqlCommand("SELECT UserId, NoPaket FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl WHERE NoPeserta=@np ORDER BY TimeInput DESC", conn)
                cmdUserList.Parameters.AddWithValue("@np", noPesertaDb) |> ignore
                use rdrUserList = cmdUserList.ExecuteReader()
                while rdrUserList.Read() do
                    let u = if rdrUserList.IsDBNull(0) then "" else rdrUserList.GetString(0)
                    let p = if rdrUserList.IsDBNull(1) then 0 else Convert.ToInt32(rdrUserList.GetValue(1))
                    if not (String.IsNullOrWhiteSpace(u)) && p > 0 then
                        altUserIds.Add((u, p))
                rdrUserList.Close()

            let psiko = ResizeArray<obj>()
            let psikoDetail = ResizeArray<obj>()
            let psikoGroupMeta = ResizeArray<obj>()
            let hasCandidate = (not (String.IsNullOrWhiteSpace(userId)) && noPaket > 0) || altUserIds.Count > 0
            if hasCandidate then
                let mutable chosenUserId = userId
                let mutable chosenNoPaket = noPaket
                if (String.IsNullOrWhiteSpace(chosenUserId) || chosenNoPaket <= 0) && altUserIds.Count > 0 then
                    let (u0, p0) = altUserIds.[0]
                    chosenUserId <- u0
                    chosenNoPaket <- p0

                let getCount (u: string) (p: int) =
                    use cmdCount = new Microsoft.Data.SqlClient.SqlCommand("SELECT COUNT(1) FROM WISECON_PSIKOTEST.dbo.TR_Psikotest WHERE UserId=@u AND NoPaket=@p", conn)
                    cmdCount.Parameters.AddWithValue("@u", u) |> ignore
                    cmdCount.Parameters.AddWithValue("@p", p) |> ignore
                    let obj = cmdCount.ExecuteScalar()
                    if isNull obj then 0 else Convert.ToInt32(obj)

                let mutable rowCount = if String.IsNullOrWhiteSpace(chosenUserId) || chosenNoPaket <= 0 then 0 else getCount chosenUserId chosenNoPaket
                if rowCount = 0 && altUserIds.Count > 0 then
                    for (u, p) in altUserIds do
                        if rowCount = 0 then
                            let c = getCount u p
                            if c > 0 then
                                chosenUserId <- u
                                chosenNoPaket <- p
                                rowCount <- c

                use cmdPs = new Microsoft.Data.SqlClient.SqlCommand("SELECT R.GroupSoal, ISNULL(R.NilaiStandard, G.NilaiStandar) AS NilaiStandard, R.NilaiGroupResult, G.NamaGroup, (SELECT SUM(mx) FROM (SELECT MAX(ISNULL(J.NoJawabanBenar,0)) mx FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtlJawaban J WHERE J.NoPaket=@p AND J.NoGroup = G.NoGroup GROUP BY J.NoUrut) X) AS NilaiMax FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult R LEFT JOIN WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup G ON G.NoPaket=@p AND G.NoGroup = TRY_CONVERT(int, R.GroupSoal) WHERE (R.NoPeserta=@np OR R.UserId=@u) ORDER BY R.GroupSoal", conn)
                cmdPs.Parameters.AddWithValue("@p", chosenNoPaket) |> ignore
                cmdPs.Parameters.AddWithValue("@u", chosenUserId) |> ignore
                cmdPs.Parameters.AddWithValue("@np", noPesertaDb) |> ignore
                use rdrPs = cmdPs.ExecuteReader()
                while rdrPs.Read() do
                    let noGroup = if rdrPs.IsDBNull(0) then "" else rdrPs.GetString(0)
                    let standar = if rdrPs.IsDBNull(1) then 0 else Convert.ToInt32(rdrPs.GetValue(1))
                    let hasil = if rdrPs.IsDBNull(2) then 0 else Convert.ToInt32(rdrPs.GetValue(2))
                    let namaGroup = if rdrPs.IsDBNull(3) then "" else rdrPs.GetString(3)
                    let nilaiMax = if rdrPs.IsDBNull(4) then 0 else Convert.ToInt32(rdrPs.GetValue(4))
                    psiko.Add(box {| noGroup = noGroup; namaGroup = namaGroup; nilaiStandard = standar; nilaiGroupResult = hasil; nilaiMax = nilaiMax |})
                rdrPs.Close()

                use cmdMeta = new Microsoft.Data.SqlClient.SqlCommand("SELECT NoGroup, NamaGroup, NilaiStandar, IsPrioritas FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup WHERE NoPaket=@p ORDER BY NoGroup", conn)
                cmdMeta.Parameters.AddWithValue("@p", chosenNoPaket) |> ignore
                use rdrMeta = cmdMeta.ExecuteReader()
                while rdrMeta.Read() do
                    let noGroup = if rdrMeta.IsDBNull(0) then 0 else Convert.ToInt32(rdrMeta.GetValue(0))
                    let namaGroup = if rdrMeta.IsDBNull(1) then "" else rdrMeta.GetString(1)
                    let nilaiStandar = if rdrMeta.IsDBNull(2) then 0 else Convert.ToInt32(rdrMeta.GetValue(2))
                    let isPrioritas = if rdrMeta.IsDBNull(3) then false else rdrMeta.GetBoolean(3)
                    psikoGroupMeta.Add(box {| noGroup = noGroup; namaGroup = namaGroup; nilaiStandar = nilaiStandar; isPrioritas = isPrioritas |})
                rdrMeta.Close()

                use cmdDtl = new Microsoft.Data.SqlClient.SqlCommand("SELECT A.NoGroup, G.NamaGroup, A.NoUrut, D.Judul, D.Deskripsi, A.JawabanDiPilih, J.Jawaban AS JawabanDipilih, (SELECT TOP 1 JJ.Jawaban FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtlJawaban JJ WHERE JJ.NoPaket=A.NoPaket AND JJ.NoGroup=A.NoGroup AND JJ.NoUrut=A.NoUrut ORDER BY ISNULL(JJ.NoJawabanBenar,0) DESC, JJ.NoJawaban) AS JawabanBenar, ISNULL(J.NoJawabanBenar,0) AS Poin FROM WISECON_PSIKOTEST.dbo.TR_Psikotest A JOIN WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtl D ON D.NoPaket=A.NoPaket AND D.NoGroup=A.NoGroup AND D.NoUrut=A.NoUrut JOIN WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup G ON G.NoPaket=A.NoPaket AND G.NoGroup=A.NoGroup LEFT JOIN WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtlJawaban J ON J.NoPaket=A.NoPaket AND J.NoGroup=A.NoGroup AND J.NoUrut=A.NoUrut AND J.NoJawaban=A.JawabanDiPilih WHERE A.UserId=@u AND A.NoPaket=@p ORDER BY A.NoGroup, A.NoUrut", conn)
                cmdDtl.Parameters.AddWithValue("@u", chosenUserId) |> ignore
                cmdDtl.Parameters.AddWithValue("@p", chosenNoPaket) |> ignore
                use rdrDtl = cmdDtl.ExecuteReader()
                while rdrDtl.Read() do
                    let noGroup = if rdrDtl.IsDBNull(0) then 0 else Convert.ToInt32(rdrDtl.GetValue(0))
                    let namaGroup = if rdrDtl.IsDBNull(1) then "" else rdrDtl.GetString(1)
                    let noUrut = if rdrDtl.IsDBNull(2) then 0 else Convert.ToInt32(rdrDtl.GetValue(2))
                    let judul = if rdrDtl.IsDBNull(3) then "" else rdrDtl.GetString(3)
                    let deskripsi = if rdrDtl.IsDBNull(4) then "" else rdrDtl.GetString(4)
                    let jawabanDipilih = if rdrDtl.IsDBNull(6) then "" else rdrDtl.GetString(6)
                    let jawabanBenar = if rdrDtl.IsDBNull(7) then "" else rdrDtl.GetString(7)
                    let poin = if rdrDtl.IsDBNull(8) then 0 else Convert.ToInt32(rdrDtl.GetValue(8))
                    psikoDetail.Add(box {| noGroup = noGroup; namaGroup = namaGroup; noUrut = noUrut; judul = judul; deskripsi = deskripsi; jawabanDipilih = jawabanDipilih; poin = poin; jawabanBenar = jawabanBenar; poinBenar = if poin > 0 then 1 else 0 |})
                rdrDtl.Close()

            this.ViewData.["BiodataJson"] <- JsonSerializer.Serialize(biodataObj, options)
            this.ViewData.["PersonalJson"] <- JsonSerializer.Serialize(personal, options)
            this.ViewData.["ExperienceJson"] <- JsonSerializer.Serialize(experience, options)
            this.ViewData.["PersonalityJson"] <- JsonSerializer.Serialize(personality, options)
            this.ViewData.["PersonalResultJson"] <- JsonSerializer.Serialize(personalRes, options)
            this.ViewData.["ExperienceResultJson"] <- JsonSerializer.Serialize(experienceRes, options)
            this.ViewData.["PersonalityResultJson"] <- JsonSerializer.Serialize(personalityRes, options)
            this.ViewData.["StatusJson"] <- JsonSerializer.Serialize(statusObj, options)
            this.ViewData.["PsikotestJson"] <- JsonSerializer.Serialize(psiko, options)
            this.ViewData.["PsikotestDetailJson"] <- JsonSerializer.Serialize(psikoDetail, options)
            this.ViewData.["PsikotestGroupMetaJson"] <- JsonSerializer.Serialize(psikoGroupMeta, options)
            this.View()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Screening/Biodata/Interview/Save")>]
    member this.SaveInterview([<FromForm>] seqNo: Nullable<int>, [<FromForm>] hasilInterview: string, [<FromForm>] remarks: string, [<FromForm>] strengthWeakness: string, [<FromForm>] likeDislike: string, [<FromForm>] answers: string, [<FromForm>] file: IFormFile) : IActionResult =
        if not seqNo.HasValue then this.BadRequest(box {| error = "Missing seqNo" |}) :> IActionResult
        else
            let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
            let options = JsonSerializerOptions(PropertyNameCaseInsensitive = true)
            let payload =
                if String.IsNullOrWhiteSpace(answers) then { personal = [||]; experience = [||]; personality = [||] }
                else JsonSerializer.Deserialize<InterviewSavePayload>(answers, options)
            let user = this.getUserName()
            let folder = @"C:\dct_docs\WISECON_PSIKOTEST\Screening\Interview"
            let newFile = this.saveUpload(file, folder)
            conn.Open()
            try
                use tran = conn.BeginTransaction()
                try
                    let sid = seqNo.Value
                    let del1 = new Microsoft.Data.SqlClient.SqlCommand("DELETE FROM WISECON_PSIKOTEST.dbo.REC_ResultPersonalInterview WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                    del1.Parameters.AddWithValue("@id", sid) |> ignore
                    del1.ExecuteNonQuery() |> ignore
                    let del2 = new Microsoft.Data.SqlClient.SqlCommand("DELETE FROM WISECON_PSIKOTEST.dbo.REC_ResultExperienceInterview WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                    del2.Parameters.AddWithValue("@id", sid) |> ignore
                    del2.ExecuteNonQuery() |> ignore
                    let del3 = new Microsoft.Data.SqlClient.SqlCommand("DELETE FROM WISECON_PSIKOTEST.dbo.REC_ResultPersonalityTest WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                    del3.Parameters.AddWithValue("@id", sid) |> ignore
                    del3.ExecuteNonQuery() |> ignore

                    for item in payload.personal do
                        let q = if isNull item.question then "" else item.question
                        let a = if isNull item.answer then "" else item.answer
                        let rmk = if isNull item.remark then "" else item.remark
                        use ins = new Microsoft.Data.SqlClient.SqlCommand("INSERT INTO WISECON_PSIKOTEST.dbo.REC_ResultPersonalInterview (JobVacancySubmittedSeqNo, MsSeqNo, Question, Answer, Remark, UserInput, TimeInput) VALUES (@id, @ms, @q, @a, @r, @u, GETDATE())", conn, tran)
                        ins.Parameters.AddWithValue("@id", sid) |> ignore
                        ins.Parameters.AddWithValue("@ms", item.msSeqNo) |> ignore
                        ins.Parameters.AddWithValue("@q", q) |> ignore
                        ins.Parameters.AddWithValue("@a", a) |> ignore
                        ins.Parameters.AddWithValue("@r", rmk) |> ignore
                        ins.Parameters.AddWithValue("@u", user) |> ignore
                        ins.ExecuteNonQuery() |> ignore

                    for item in payload.experience do
                        let q = if isNull item.question then "" else item.question
                        let a = if isNull item.answer then "" else item.answer
                        let rmk = if isNull item.remark then "" else item.remark
                        use ins = new Microsoft.Data.SqlClient.SqlCommand("INSERT INTO WISECON_PSIKOTEST.dbo.REC_ResultExperienceInterview (JobVacancySubmittedSeqNo, MsSeqNo, Question, Answer, Remark, UserInput, TimeInput) VALUES (@id, @ms, @q, @a, @r, @u, GETDATE())", conn, tran)
                        ins.Parameters.AddWithValue("@id", sid) |> ignore
                        ins.Parameters.AddWithValue("@ms", item.msSeqNo) |> ignore
                        ins.Parameters.AddWithValue("@q", q) |> ignore
                        ins.Parameters.AddWithValue("@a", a) |> ignore
                        ins.Parameters.AddWithValue("@r", rmk) |> ignore
                        ins.Parameters.AddWithValue("@u", user) |> ignore
                        ins.ExecuteNonQuery() |> ignore

                    for item in payload.personality do
                        let a = if isNull item.aspek then "" else item.aspek
                        let u = if isNull item.uraian then "" else item.uraian
                        use ins = new Microsoft.Data.SqlClient.SqlCommand("INSERT INTO WISECON_PSIKOTEST.dbo.REC_ResultPersonalityTest (JobVacancySubmittedSeqNo, MsSeqNo, Aspek, Uraian, Nilai, UserInput, TimeInput) VALUES (@id, @ms, @a, @u, @n, @user, GETDATE())", conn, tran)
                        ins.Parameters.AddWithValue("@id", sid) |> ignore
                        ins.Parameters.AddWithValue("@ms", item.msSeqNo) |> ignore
                        ins.Parameters.AddWithValue("@a", a) |> ignore
                        ins.Parameters.AddWithValue("@u", u) |> ignore
                        ins.Parameters.AddWithValue("@n", item.nilai) |> ignore
                        ins.Parameters.AddWithValue("@user", user) |> ignore
                        ins.ExecuteNonQuery() |> ignore

                    use chk = new Microsoft.Data.SqlClient.SqlCommand("SELECT InterviewResultFileName FROM WISECON_PSIKOTEST.dbo.REC_ScreeningStatus WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                    chk.Parameters.AddWithValue("@id", sid) |> ignore
                    let existingFileObj = chk.ExecuteScalar()
                    let existingFile = if isNull existingFileObj then "" else existingFileObj.ToString()
                    let finalFile = if String.IsNullOrWhiteSpace(newFile) then existingFile else newFile
                    let statusScreening = if String.IsNullOrWhiteSpace(hasilInterview) then "Belum Proses" else "Interview Proses"

                    let countCmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT COUNT(1) FROM WISECON_PSIKOTEST.dbo.REC_ScreeningStatus WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                    countCmd.Parameters.AddWithValue("@id", sid) |> ignore
                    let count = Convert.ToInt32(countCmd.ExecuteScalar())
                    if count > 0 then
                        use upd = new Microsoft.Data.SqlClient.SqlCommand("UPDATE WISECON_PSIKOTEST.dbo.REC_ScreeningStatus SET ResultInterviewTes=@hasil, StatusTesInterview=@hasil, InterviewResultRemarks=@remarks, InterviewResultFileName=@file, InterviewStrengthWeakness=@strength, InterviewLikeDislike=@likeDislike, StatusScreening=@screening, UserEdit=@u, TimeEdit=GETDATE() WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                        upd.Parameters.AddWithValue("@hasil", (if isNull hasilInterview then "" else hasilInterview)) |> ignore
                        upd.Parameters.AddWithValue("@remarks", (if isNull remarks then "" else remarks)) |> ignore
                        upd.Parameters.AddWithValue("@file", (if isNull finalFile then "" else finalFile)) |> ignore
                        upd.Parameters.AddWithValue("@strength", (if isNull strengthWeakness then "" else strengthWeakness)) |> ignore
                        upd.Parameters.AddWithValue("@likeDislike", (if isNull likeDislike then "" else likeDislike)) |> ignore
                        upd.Parameters.AddWithValue("@screening", statusScreening) |> ignore
                        upd.Parameters.AddWithValue("@u", user) |> ignore
                        upd.Parameters.AddWithValue("@id", sid) |> ignore
                        upd.ExecuteNonQuery() |> ignore
                    else
                        use ins = new Microsoft.Data.SqlClient.SqlCommand("INSERT INTO WISECON_PSIKOTEST.dbo.REC_ScreeningStatus (JobVacancySubmittedSeqNo, ResultInterviewTes, StatusTesInterview, InterviewResultRemarks, InterviewResultFileName, InterviewStrengthWeakness, InterviewLikeDislike, StatusScreening, UserInput, TimeInput, UserEdit, TimeEdit) VALUES (@id, @hasil, @hasil, @remarks, @file, @strength, @likeDislike, @screening, @u, GETDATE(), @u, GETDATE())", conn, tran)
                        ins.Parameters.AddWithValue("@id", sid) |> ignore
                        ins.Parameters.AddWithValue("@hasil", (if isNull hasilInterview then "" else hasilInterview)) |> ignore
                        ins.Parameters.AddWithValue("@remarks", (if isNull remarks then "" else remarks)) |> ignore
                        ins.Parameters.AddWithValue("@file", (if isNull finalFile then "" else finalFile)) |> ignore
                        ins.Parameters.AddWithValue("@strength", (if isNull strengthWeakness then "" else strengthWeakness)) |> ignore
                        ins.Parameters.AddWithValue("@likeDislike", (if isNull likeDislike then "" else likeDislike)) |> ignore
                        ins.Parameters.AddWithValue("@screening", statusScreening) |> ignore
                        ins.Parameters.AddWithValue("@u", user) |> ignore
                        ins.ExecuteNonQuery() |> ignore

                    tran.Commit()
                    this.Ok(box {| ok = true |})
                with ex ->
                    tran.Rollback()
                    this.BadRequest(box {| error = ex.Message |})
            finally
                conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Screening/Biodata/Psikotest/Save")>]
    member this.SavePsikotest([<FromForm>] seqNo: Nullable<int>, [<FromForm>] hasilPsikotest: string, [<FromForm>] file: IFormFile) : IActionResult =
        if not seqNo.HasValue then this.BadRequest(box {| error = "Missing seqNo" |}) :> IActionResult
        else
            let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
            let user = this.getUserName()
            let folder = @"C:\dct_docs\WISECON_PSIKOTEST\Screening\Psikotest"
            let newFile = this.saveUpload(file, folder)
            conn.Open()
            try
                use tran = conn.BeginTransaction()
                try
                    let sid = seqNo.Value
                    use chk = new Microsoft.Data.SqlClient.SqlCommand("SELECT PsikotesResultFileName FROM WISECON_PSIKOTEST.dbo.REC_ScreeningStatus WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                    chk.Parameters.AddWithValue("@id", sid) |> ignore
                    let existingFileObj = chk.ExecuteScalar()
                    let existingFile = if isNull existingFileObj then "" else existingFileObj.ToString()
                    let finalFile = if String.IsNullOrWhiteSpace(newFile) then existingFile else newFile
                    let statusScreening = if String.IsNullOrWhiteSpace(hasilPsikotest) then "Belum Proses" else "Psikotest Proses"

                    let countCmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT COUNT(1) FROM WISECON_PSIKOTEST.dbo.REC_ScreeningStatus WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                    countCmd.Parameters.AddWithValue("@id", sid) |> ignore
                    let count = Convert.ToInt32(countCmd.ExecuteScalar())
                    if count > 0 then
                        use upd = new Microsoft.Data.SqlClient.SqlCommand("UPDATE WISECON_PSIKOTEST.dbo.REC_ScreeningStatus SET StatusPsikotes=@hasil, StatusTesPsikotes=@hasil, PsikotesResultFileName=@file, StatusScreening=@screening, UserEdit=@u, TimeEdit=GETDATE() WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                        upd.Parameters.AddWithValue("@hasil", (if isNull hasilPsikotest then "" else hasilPsikotest)) |> ignore
                        upd.Parameters.AddWithValue("@file", (if isNull finalFile then "" else finalFile)) |> ignore
                        upd.Parameters.AddWithValue("@screening", statusScreening) |> ignore
                        upd.Parameters.AddWithValue("@u", user) |> ignore
                        upd.Parameters.AddWithValue("@id", sid) |> ignore
                        upd.ExecuteNonQuery() |> ignore
                    else
                        use ins = new Microsoft.Data.SqlClient.SqlCommand("INSERT INTO WISECON_PSIKOTEST.dbo.REC_ScreeningStatus (JobVacancySubmittedSeqNo, StatusPsikotes, StatusTesPsikotes, PsikotesResultFileName, StatusScreening, UserInput, TimeInput, UserEdit, TimeEdit) VALUES (@id, @hasil, @hasil, @file, @screening, @u, GETDATE(), @u, GETDATE())", conn, tran)
                        ins.Parameters.AddWithValue("@id", sid) |> ignore
                        ins.Parameters.AddWithValue("@hasil", (if isNull hasilPsikotest then "" else hasilPsikotest)) |> ignore
                        ins.Parameters.AddWithValue("@file", (if isNull finalFile then "" else finalFile)) |> ignore
                        ins.Parameters.AddWithValue("@screening", statusScreening) |> ignore
                        ins.Parameters.AddWithValue("@u", user) |> ignore
                        ins.ExecuteNonQuery() |> ignore
                    tran.Commit()
                    this.Ok(box {| ok = true |})
                with ex ->
                    tran.Rollback()
                    this.BadRequest(box {| error = ex.Message |})
            finally
                conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Screening/Biodata/Approve")>]
    member this.Approve([<FromBody>] req: ApproveRequest) : IActionResult =
        if obj.ReferenceEquals(req, null) || isNull req.items || req.items.Length = 0 then this.BadRequest(box {| error = "No data" |}) :> IActionResult
        else
            let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
            let user = this.getUserName()
            conn.Open()
            try
                let warn = ResizeArray<string>()
                for item in req.items do
                    let sid = if isNull item.kodeBiodata then 0 else Int32.Parse(item.kodeBiodata)
                    use cmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT StatusScreening, ResultInterviewTes, StatusPsikotes, StatusVerif FROM WISECON_PSIKOTEST.dbo.VW_REC_Screening WHERE SeqNo=@id", conn)
                    cmd.Parameters.AddWithValue("@id", sid) |> ignore
                    use rdr = cmd.ExecuteReader()
                    let mutable statusScreening = ""
                    let mutable hasilInterview = ""
                    let mutable hasilPsikotes = ""
                    let mutable statusVerif = ""
                    if rdr.Read() then
                        statusScreening <- if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                        hasilInterview <- if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                        hasilPsikotes <- if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                        statusVerif <- if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                    rdr.Close()
                    if String.IsNullOrWhiteSpace(statusScreening) || statusScreening = "Belum Proses" then
                        warn.Add(string sid)
                    else
                        let mutable newVerif = statusVerif
                        if String.Equals(hasilInterview, "Tidak Lulus", StringComparison.OrdinalIgnoreCase) || String.Equals(hasilPsikotes, "Tidak Lulus", StringComparison.OrdinalIgnoreCase) then
                            newVerif <- "Tidak Lulus"
                        elif String.Equals(hasilInterview, "Lulus", StringComparison.OrdinalIgnoreCase) && String.Equals(hasilPsikotes, "Lulus", StringComparison.OrdinalIgnoreCase) then
                            newVerif <- "Lulus"
                        let existsCmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT COUNT(1) FROM WISECON_PSIKOTEST.dbo.REC_ScreeningStatus WHERE JobVacancySubmittedSeqNo=@id", conn)
                        existsCmd.Parameters.AddWithValue("@id", sid) |> ignore
                        let count = Convert.ToInt32(existsCmd.ExecuteScalar())
                        if count > 0 then
                            use upd = new Microsoft.Data.SqlClient.SqlCommand("UPDATE WISECON_PSIKOTEST.dbo.REC_ScreeningStatus SET StatusVerif=@v, UserVerif=@u, TimeEdit=GETDATE() WHERE JobVacancySubmittedSeqNo=@id", conn)
                            upd.Parameters.AddWithValue("@v", (if isNull newVerif then "" else newVerif)) |> ignore
                            upd.Parameters.AddWithValue("@u", user) |> ignore
                            upd.Parameters.AddWithValue("@id", sid) |> ignore
                            upd.ExecuteNonQuery() |> ignore
                        else
                            use ins = new Microsoft.Data.SqlClient.SqlCommand("INSERT INTO WISECON_PSIKOTEST.dbo.REC_ScreeningStatus (JobVacancySubmittedSeqNo, StatusVerif, UserVerif, UserInput, TimeInput, UserEdit, TimeEdit) VALUES (@id, @v, @u, @u, GETDATE(), @u, GETDATE())", conn)
                            ins.Parameters.AddWithValue("@id", sid) |> ignore
                            ins.Parameters.AddWithValue("@v", (if isNull newVerif then "" else newVerif)) |> ignore
                            ins.Parameters.AddWithValue("@u", user) |> ignore
                            ins.ExecuteNonQuery() |> ignore
                let msg =
                    if warn.Count > 0 then
                        "Status screening belum diproses untuk kode biodata: " + String.Join(", ", warn)
                    else ""
                this.Ok(box {| ok = true; message = msg |})
            finally
                conn.Close()
