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
type PersonalityItem = { msSeqNo: int; aspek: string; uraian: string; nilai: int }

[<CLIMutable>]
type InterviewSavePayload = { personal: InterviewItem array; experience: InterviewItem array; personality: PersonalityItem array }

[<ApiController>]
type ScreeningController (db: IDbConnection) =
    inherit Controller()

    member private this.getQueryValue(key: string) =
        if this.Request.Query.ContainsKey(key) then this.Request.Query.[key].ToString() else ""

    member private this.applyFilters(cmd: Microsoft.Data.SqlClient.SqlCommand) =
        let mutable whereClause = " WHERE 1=1 "
        let fromDateStr = this.getQueryValue("fromDate")
        if not (String.IsNullOrWhiteSpace(fromDateStr)) then
            match DateTime.TryParse(fromDateStr) with
            | true, v ->
                whereClause <- whereClause + " AND TimeInput >= @fromDate "
                cmd.Parameters.AddWithValue("@fromDate", v.Date) |> ignore
            | _ -> ()
        let toDateStr = this.getQueryValue("toDate")
        if not (String.IsNullOrWhiteSpace(toDateStr)) then
            match DateTime.TryParse(toDateStr) with
            | true, v ->
                whereClause <- whereClause + " AND TimeInput <= @toDate "
                cmd.Parameters.AddWithValue("@toDate", v.Date.AddDays(1.0).AddSeconds(-1.0)) |> ignore
            | _ -> ()
        let statusSeleksi = this.getQueryValue("statusSeleksi")
        if not (String.IsNullOrWhiteSpace(statusSeleksi)) then
            whereClause <- whereClause + " AND StatusScreening = @statusSeleksi "
            cmd.Parameters.AddWithValue("@statusSeleksi", statusSeleksi) |> ignore
        let statusVerifikasi = this.getQueryValue("statusVerifikasi")
        if not (String.IsNullOrWhiteSpace(statusVerifikasi)) then
            whereClause <- whereClause + " AND StatusVerif = @statusVerifikasi "
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
            "SELECT JobCode, JobPosition, " +
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
            "GROUP BY JobCode, JobPosition " +
            "ORDER BY JobCode"
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
                rows.Add(box {| code = code; name = name; belumProses = belum; interviewProses = intProses; interviewTidakDatang = intTidak; psikotestProses = psiProses; psikotestTidakDatang = psiTidak; statusLulus = lulus; statusTidakLulus = tdk; timeInput = timeInput |})
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
        let whereClause = this.applyFilters(cmd)
        cmd.CommandText <-
            "SELECT JobCode, JobPosition, Name, StatusScreening, StatusVerif, ResultInterviewTes, StatusPsikotes, Sex, UserEmail, PhoneNo, PendidikanTerakhir, TimeInput " +
            "FROM WISECON_PSIKOTEST.dbo.VW_REC_Screening " +
            whereClause +
            "ORDER BY JobCode, Name"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            use wb = new XLWorkbook()
            let ws = wb.AddWorksheet("Screening")
            ws.Cell(1,1).Value <- "JobCode"
            ws.Cell(1,2).Value <- "JobPosition"
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
                ws.Cell(row,12).Value <- if rdr.IsDBNull(11) then DateTime.MinValue else rdr.GetDateTime(11)
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
        let whereClause = this.applyFilters(cmd)
        let kodeLowongan = this.getQueryValue("kodeLowongan")
        if not (String.IsNullOrWhiteSpace(kodeLowongan)) then
            cmd.Parameters.AddWithValue("@kodeLowongan", kodeLowongan) |> ignore
        cmd.CommandText <-
            "SELECT SeqNo, JobCode, JobPosition, BatchNo, Name, StatusScreening, StatusVerif, ResultInterviewTes, StatusPsikotes, Sex, UserEmail, PhoneNo, PendidikanTerakhir, CVFileName " +
            "FROM WISECON_PSIKOTEST.dbo.VW_REC_Screening " +
            whereClause +
            (if String.IsNullOrWhiteSpace(kodeLowongan) then "" else " AND JobCode = @kodeLowongan ") +
            "ORDER BY Name"
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
                rows.Add(box {| kodeLowongan = jobCode; kodeBiodata = string seqNo; posLowongan = jobPos; batch = batch; name = name; statusScreening = statusScreening; statusVerifikasi = statusVerif; hasilInterview = hasilInterview; hasilPsikotest = hasilPsikotest; sex = sex; email = email; phone = phone; pendidikanTerakhir = pendidikan; cvFileName = cv |})
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
    member this.Paper([<FromQuery>] kodeBiodata: int) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        let options = JsonSerializerOptions(PropertyNamingPolicy = JsonNamingPolicy.CamelCase)
        options.PropertyNameCaseInsensitive <- true
        conn.Open()
        try
            use cmdBio = new Microsoft.Data.SqlClient.SqlCommand("SELECT SeqNo, Name, UserEmail, JobPosition, PhoneNo, PendidikanTerakhir, StatusScreening, StatusVerif, ResultInterviewTes, StatusPsikotes FROM WISECON_PSIKOTEST.dbo.VW_REC_Screening WHERE SeqNo=@id", conn)
            cmdBio.Parameters.AddWithValue("@id", kodeBiodata) |> ignore
            use rdrBio = cmdBio.ExecuteReader()
            let mutable biodataObj = box {| seqNo = kodeBiodata; name = ""; email = ""; job = ""; phone = ""; pendidikan = ""; statusScreening = ""; statusVerif = ""; hasilInterview = ""; hasilPsikotest = "" |}
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
                biodataObj <- box {| seqNo = kodeBiodata; name = name; email = email; job = job; phone = phone; pendidikan = pendidikan; statusScreening = statusScreening; statusVerif = statusVerif; hasilInterview = hasilInterview; hasilPsikotest = hasilPsikotest |}
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
                cmd.Parameters.AddWithValue("@id", kodeBiodata) |> ignore
                use rdr = cmd.ExecuteReader()
                let rows = ResizeArray<obj>()
                while rdr.Read() do
                    rows.Add(mapFn rdr)
                rdr.Close()
                rows

            let personal = loadList "SELECT SeqNo, Question FROM WISECON_PSIKOTEST.dbo.REC_MsPersonalInterview ORDER BY SeqNo" (fun r ->
                let ms = Convert.ToInt32(r.GetInt64(0))
                let q = if r.IsDBNull(1) then "" else r.GetString(1)
                box {| msSeqNo = ms; question = q |})
            let experience = loadList "SELECT SeqNo, Question FROM WISECON_PSIKOTEST.dbo.REC_MsExperienceInterview ORDER BY SeqNo" (fun r ->
                let ms = Convert.ToInt32(r.GetInt64(0))
                let q = if r.IsDBNull(1) then "" else r.GetString(1)
                box {| msSeqNo = ms; question = q |})
            let personality = loadList "SELECT SeqNo, Aspek, Uraian FROM WISECON_PSIKOTEST.dbo.REC_MsPersonalityTest ORDER BY SeqNo" (fun r ->
                let ms = Convert.ToInt32(r.GetInt64(0))
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
                let n = if r.IsDBNull(1) then 0 else r.GetInt32(1)
                box {| msSeqNo = ms; nilai = n |})

            let statusCmd = new Microsoft.Data.SqlClient.SqlCommand("SELECT ResultInterviewTes, InterviewResultRemarks, InterviewResultFileName, StatusPsikotes, PsikotesResultFileName FROM WISECON_PSIKOTEST.dbo.REC_ScreeningStatus WHERE JobVacancySubmittedSeqNo=@id", conn)
            statusCmd.Parameters.AddWithValue("@id", kodeBiodata) |> ignore
            use statusR = statusCmd.ExecuteReader()
            let mutable statusObj = box {| hasilInterview = ""; remarks = ""; interviewFile = ""; hasilPsikotest = ""; psikotestFile = "" |}
            if statusR.Read() then
                let hasilInterview = if statusR.IsDBNull(0) then "" else statusR.GetString(0)
                let remarks = if statusR.IsDBNull(1) then "" else statusR.GetString(1)
                let interviewFile = if statusR.IsDBNull(2) then "" else statusR.GetString(2)
                let hasilPsikotest = if statusR.IsDBNull(3) then "" else statusR.GetString(3)
                let psikotestFile = if statusR.IsDBNull(4) then "" else statusR.GetString(4)
                statusObj <- box {| hasilInterview = hasilInterview; remarks = remarks; interviewFile = interviewFile; hasilPsikotest = hasilPsikotest; psikotestFile = psikotestFile |}
            statusR.Close()

            let psiko = ResizeArray<obj>()
            use cmdPs = new Microsoft.Data.SqlClient.SqlCommand("SELECT NoGroup, NilaiStandard, NilaiGroupResult FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult WHERE UserId = (SELECT TOP 1 UserId FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl WHERE NoPeserta = (SELECT NoPeserta FROM WISECON_PSIKOTEST.dbo.MS_Peserta WHERE ID=@id)) ORDER BY NoGroup", conn)
            cmdPs.Parameters.AddWithValue("@id", kodeBiodata) |> ignore
            use rdrPs = cmdPs.ExecuteReader()
            while rdrPs.Read() do
                let noGroup = if rdrPs.IsDBNull(0) then 0 else rdrPs.GetInt32(0)
                let standar = if rdrPs.IsDBNull(1) then 0 else rdrPs.GetInt32(1)
                let hasil = if rdrPs.IsDBNull(2) then 0 else rdrPs.GetInt32(2)
                psiko.Add(box {| noGroup = noGroup; nilaiStandard = standar; nilaiGroupResult = hasil |})
            rdrPs.Close()

            this.ViewData.["BiodataJson"] <- JsonSerializer.Serialize(biodataObj, options)
            this.ViewData.["PersonalJson"] <- JsonSerializer.Serialize(personal, options)
            this.ViewData.["ExperienceJson"] <- JsonSerializer.Serialize(experience, options)
            this.ViewData.["PersonalityJson"] <- JsonSerializer.Serialize(personality, options)
            this.ViewData.["PersonalResultJson"] <- JsonSerializer.Serialize(personalRes, options)
            this.ViewData.["ExperienceResultJson"] <- JsonSerializer.Serialize(experienceRes, options)
            this.ViewData.["PersonalityResultJson"] <- JsonSerializer.Serialize(personalityRes, options)
            this.ViewData.["StatusJson"] <- JsonSerializer.Serialize(statusObj, options)
            this.ViewData.["PsikotestJson"] <- JsonSerializer.Serialize(psiko, options)
            this.View()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Screening/Biodata/Interview/Save")>]
    member this.SaveInterview([<FromForm>] seqNo: Nullable<int>, [<FromForm>] hasilInterview: string, [<FromForm>] remarks: string, [<FromForm>] answers: string, [<FromForm>] file: IFormFile) : IActionResult =
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
                        use upd = new Microsoft.Data.SqlClient.SqlCommand("UPDATE WISECON_PSIKOTEST.dbo.REC_ScreeningStatus SET ResultInterviewTes=@hasil, StatusTesInterview=@hasil, InterviewResultRemarks=@remarks, InterviewResultFileName=@file, StatusScreening=@screening, UserEdit=@u, TimeEdit=GETDATE() WHERE JobVacancySubmittedSeqNo=@id", conn, tran)
                        upd.Parameters.AddWithValue("@hasil", (if isNull hasilInterview then "" else hasilInterview)) |> ignore
                        upd.Parameters.AddWithValue("@remarks", (if isNull remarks then "" else remarks)) |> ignore
                        upd.Parameters.AddWithValue("@file", (if isNull finalFile then "" else finalFile)) |> ignore
                        upd.Parameters.AddWithValue("@screening", statusScreening) |> ignore
                        upd.Parameters.AddWithValue("@u", user) |> ignore
                        upd.Parameters.AddWithValue("@id", sid) |> ignore
                        upd.ExecuteNonQuery() |> ignore
                    else
                        use ins = new Microsoft.Data.SqlClient.SqlCommand("INSERT INTO WISECON_PSIKOTEST.dbo.REC_ScreeningStatus (JobVacancySubmittedSeqNo, ResultInterviewTes, StatusTesInterview, InterviewResultRemarks, InterviewResultFileName, StatusScreening, UserInput, TimeInput, UserEdit, TimeEdit) VALUES (@id, @hasil, @hasil, @remarks, @file, @screening, @u, GETDATE(), @u, GETDATE())", conn, tran)
                        ins.Parameters.AddWithValue("@id", sid) |> ignore
                        ins.Parameters.AddWithValue("@hasil", (if isNull hasilInterview then "" else hasilInterview)) |> ignore
                        ins.Parameters.AddWithValue("@remarks", (if isNull remarks then "" else remarks)) |> ignore
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
