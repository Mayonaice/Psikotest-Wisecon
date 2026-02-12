namespace PsikotestWisesa.Controllers

open System
open System.IO
open System.Data
open System.Text.RegularExpressions
open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authorization

[<CLIMutable>]
type PesertaRequest = {
    NoKTP: string
    NamaPeserta: string
    Alamat: string
    NoHP: string
    Email: string
    JenisKelamin: string
    JobName: string
    LastEducation: string
    AttachmentBase64: string
    UserInput: string
}

[<ApiController>]
type ApplicationsController (db: IDbConnection) =
    inherit Controller()

    [<AllowAnonymous>]
    [<HttpPost>]
    [<Route("Applications/Peserta")>]
    member this.CreatePeserta ([<FromBody>] req: PesertaRequest) : IActionResult =
        let expected = "PS1K0T35T_W1S3C0N@1312245"
        let auth = if this.Request.Headers.ContainsKey("Authorization") then this.Request.Headers.["Authorization"].ToString() else null
        let xauth = if this.Request.Headers.ContainsKey("X-Auth-Token") then this.Request.Headers.["X-Auth-Token"].ToString() else null
        let qauth = if this.Request.Query.ContainsKey("token") then this.Request.Query.["token"].ToString() else null
        let mutable ok = false
        if not (isNull auth) && auth.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase) then
            let t = auth.Substring(7)
            ok <- t = expected
        elif not (isNull auth) then
            ok <- auth = expected
        elif not (isNull xauth) then
            ok <- xauth = expected
        elif not (isNull qauth) then
            ok <- qauth = expected
        if not ok then this.Unauthorized() :> IActionResult
        else if String.IsNullOrWhiteSpace(req.UserInput) then this.BadRequest(box {| error = "UserInput is required" |}) :> IActionResult
        else
        let dir = "C:\\dct_docs\\WISECON_PSIKOTEST\\Resume\\"
        Directory.CreateDirectory(dir) |> ignore

        let base64 = if String.IsNullOrWhiteSpace(req.AttachmentBase64) then "" else req.AttachmentBase64
        let mutable dataPart = base64
        let mutable ext = ".pdf"
        if base64.Contains(",") then
            let parts = base64.Split(',', 2)
            let header = parts.[0]
            dataPart <- parts.[1]
            if header.StartsWith("data:") then
                let m = Regex.Match(header, "data:([^;]+)")
                if m.Success then
                    let mime = m.Groups.[1].Value.ToLowerInvariant()
                    ext <-
                        if mime.Contains("pdf") then ".pdf"
                        elif mime.Contains("png") then ".png"
                        elif mime.Contains("jpeg") || mime.Contains("jpg") then ".jpg"
                        elif mime.Contains("msword") || mime.Contains("word") then ".doc"
                        elif mime.Contains("officedocument.wordprocessingml.document") then ".docx"
                        else ".bin"

        let fileName = DateTime.Now.ToString("yyyyMMddHHmmssfff") + ext
        let fullPath = Path.Combine(dir, fileName)
        try
            let bytes = Convert.FromBase64String(dataPart)
            System.IO.File.WriteAllBytes(fullPath, bytes)
        with
        | :? FormatException -> ()

        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandText <- "dbo.WEB_SP_MsPesertaInsert"
        cmd.CommandType <- CommandType.StoredProcedure
        cmd.CommandTimeout <- 120
        cmd.Parameters.AddWithValue("@NoKTP", req.NoKTP) |> ignore
        cmd.Parameters.AddWithValue("@NamaPeserta", req.NamaPeserta) |> ignore
        cmd.Parameters.AddWithValue("@Alamat", req.Alamat) |> ignore
        cmd.Parameters.AddWithValue("@NoHP", req.NoHP) |> ignore
        cmd.Parameters.AddWithValue("@Email", req.Email) |> ignore
        cmd.Parameters.AddWithValue("@JenisKelamin", req.JenisKelamin) |> ignore
        cmd.Parameters.AddWithValue("@JobName", req.JobName) |> ignore
        cmd.Parameters.AddWithValue("@Attachment", fileName) |> ignore
        cmd.Parameters.AddWithValue("@LastEducation", req.LastEducation) |> ignore
        let userInput = req.UserInput
        cmd.Parameters.AddWithValue("@UserInput", userInput) |> ignore

        conn.Open()
        try conn.ChangeDatabase("WISECON_PSIKOTEST") with _ -> ()
        try
            let _ = cmd.ExecuteNonQuery()
            use verify = new Microsoft.Data.SqlClient.SqlCommand()
            verify.Connection <- conn
            verify.CommandText <- "SELECT TOP 1 ID, NoPeserta FROM WISECON_PSIKOTEST.dbo.MS_Peserta WHERE NoKTP=@NoKTP ORDER BY ID DESC"
            verify.Parameters.AddWithValue("@NoKTP", req.NoKTP) |> ignore
            use r = verify.ExecuteReader()
            let mutable inserted = false
            let mutable id = 0
            let mutable noPeserta = ""
            if r.Read() then
                inserted <- true
                id <- r.GetInt32(0)
                noPeserta <- r.GetString(1)
            r.Close()
            this.Created("/Applications/Peserta", box {| fileName = fileName; path = fullPath; inserted = inserted; id = id; noPeserta = noPeserta |})
        finally
            conn.Close()
