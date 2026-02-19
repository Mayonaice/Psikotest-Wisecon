namespace PsikotestWisesa.Controllers

open System
open System.Data
open System.Text
open System.Security.Cryptography
open System.Collections.Concurrent
open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authorization
open Microsoft.Extensions.Configuration
open Microsoft.Data.SqlClient

[<CLIMutable>]
type ExamLoginRequest = { Token: string; NoKTP: string; Email: string }

[<CLIMutable>]
type ExamFinishRequest = { Token: string }

[<CLIMutable>]
type ExamSubmitAnswerRequest = { Token: string; NoGroup: int; NoUrut: int; JawabanDiPilih: int }

[<CLIMutable>]
type ExamGroupStartRequest = { Token: string; NoGroup: int }

type ExamAnswer = { noJawaban: int; jawaban: string; poin: int }
type ExamQuestion = { noUrut: int; judul: string; deskripsi: string; answers: ResizeArray<ExamAnswer> }
type ExamGroup = { noGroup: int; namaGroup: string; minSoal: int; waktu: int; random: bool; isPrioritas: bool; questions: ResizeArray<ExamQuestion> }

type ExamController (db: IDbConnection, cfg: IConfiguration) =
    inherit Controller()

    static let groupStartMap = ConcurrentDictionary<string, DateTime>()

    let decrypt64 (cipherTextBase64: string) (encryptionKey: string) =
        if String.IsNullOrWhiteSpace(cipherTextBase64) || String.IsNullOrWhiteSpace(encryptionKey) then
            ""
        else
            let key8 = if encryptionKey.Length >= 8 then encryptionKey.Substring(0, 8) else ""
            if String.IsNullOrWhiteSpace(key8) then "" else
            try
                use des = new DESCryptoServiceProvider()
                let keyBytes = Encoding.UTF8.GetBytes(key8)
                let ivBytes = [| 0x12uy; 0x34uy; 0x56uy; 0x78uy; 0x90uy; 0xABuy; 0xCDuy; 0xEFuy |]
                let cipherBytes = Convert.FromBase64String(cipherTextBase64)
                use decryptor = des.CreateDecryptor(keyBytes, ivBytes)
                let plainBytes = decryptor.TransformFinalBlock(cipherBytes, 0, cipherBytes.Length)
                Encoding.UTF8.GetString(plainBytes)
            with _ ->
                ""

    let getSetting (keys: string list) =
        keys
        |> List.tryPick (fun k ->
            let v = cfg.[k]
            if String.IsNullOrWhiteSpace(v) then None else Some v)
        |> Option.defaultValue ""

    let getEncryptionKey () =
        getSetting [ "Psikotest:EncryptionKey"; "EncryptionKey" ]

    let tryParseToken (plain: string) =
        if String.IsNullOrWhiteSpace(plain) then None else
        let parts = plain.Split('&', StringSplitOptions.RemoveEmptyEntries)
        let dict =
            parts
            |> Array.choose (fun p ->
                let i = p.IndexOf('=')
                if i <= 0 then None else
                let k = p.Substring(0, i).Trim().ToUpperInvariant()
                let v = p.Substring(i + 1)
                Some(k, v))
            |> dict
        let tryGet k =
            if dict.ContainsKey(k) then dict.[k] else ""
        let userId = tryGet "U"
        let noPaketStr = tryGet "NP"
        let waktuStr = tryGet "WT"
        let noKtp = tryGet "KTP"
        let email = tryGet "EM"
        let mutable noPaket = 0L
        let mutable waktu = DateTime.MinValue
        let ok1 = Int64.TryParse(noPaketStr, &noPaket)
        let ok2 = DateTime.TryParse(waktuStr, &waktu)
        if String.IsNullOrWhiteSpace(userId) || not ok1 || not ok2 then None
        else Some(userId, noPaket, waktu, noKtp, email)

    let norm (s: string) =
        if isNull s then "" else s.Trim()

    let eqi (a: string) (b: string) =
        String.Equals(norm a, norm b, StringComparison.OrdinalIgnoreCase)

    member private this.hasFinalResult(conn: SqlConnection, userId: string) =
        use cmd = new SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT CASE WHEN EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult WHERE UserId=@u) THEN 1 ELSE 0 END"
        cmd.Parameters.AddWithValue("@u", userId) |> ignore
        let obj = cmd.ExecuteScalar()
        try Convert.ToInt32(obj) = 1 with _ -> false

    member private this.getToleranceMinutes(conn: SqlConnection, noPaket: int64) =
        use cmd = new SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT ISNULL(ToleransiWaktu,0) FROM WISECON_PSIKOTEST.dbo.MS_PaketSoal WHERE NoPaket=@p"
        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
        let obj = cmd.ExecuteScalar()
        try Convert.ToInt32(obj) with _ -> 0

    member private this.getAssignment(conn: SqlConnection, userId: string, noPaket: int64) =
        use cmd = new SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <-
            "SELECT TOP 1 D.IdNoPeserta, D.NoPeserta, D.UserId, D.NoPaket, D.WaktuTest, D.StartTest, D.TimeEdit, ISNULL(D.block,0) AS IsBlocked, " +
            "P.NoKTP, P.Email, ISNULL(PS.NamaPaket,'') AS NamaPaket " +
            "FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl D " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.MS_Peserta P ON P.ID=D.IdNoPeserta " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.MS_PaketSoal PS ON PS.NoPaket=D.NoPaket " +
            "WHERE D.UserId=@u AND D.NoPaket=@p " +
            "ORDER BY D.TimeInput DESC"
        cmd.Parameters.AddWithValue("@u", userId) |> ignore
        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
        use rdr = cmd.ExecuteReader()
        if rdr.Read() then
            let idNoPeserta = if rdr.IsDBNull(0) then 0L else Convert.ToInt64(rdr.GetValue(0))
            let noPeserta = if rdr.IsDBNull(1) then "" else Convert.ToString(rdr.GetValue(1))
            let userId2 = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
            let noPaket2 = if rdr.IsDBNull(3) then 0L else Convert.ToInt64(rdr.GetValue(3))
            let waktuTest = if rdr.IsDBNull(4) then Nullable() else Nullable(rdr.GetDateTime(4))
            let startTest = if rdr.IsDBNull(5) then Nullable() else Nullable(rdr.GetDateTime(5))
            let timeEdit = if rdr.IsDBNull(6) then Nullable() else Nullable(rdr.GetDateTime(6))
            let isBlocked = if rdr.IsDBNull(7) then false else (Convert.ToInt32(rdr.GetValue(7)) <> 0)
            let noKtp = if rdr.IsDBNull(8) then "" else rdr.GetString(8)
            let email = if rdr.IsDBNull(9) then "" else rdr.GetString(9)
            let namaPaket = if rdr.IsDBNull(10) then "" else rdr.GetString(10)
            Some(idNoPeserta, noPeserta, userId2, noPaket2, waktuTest, startTest, timeEdit, isBlocked, noKtp, email, namaPaket)
        else
            None

    member private this.getGroupDurationMinutes(conn: SqlConnection, noPaket: int64, noGroup: int) =
        use cmd = new SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT ISNULL(WaktuPengerjaan,0) FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup WHERE NoPaket=@p AND NoGroup=@g"
        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
        cmd.Parameters.AddWithValue("@g", noGroup) |> ignore
        let obj = cmd.ExecuteScalar()
        try Convert.ToInt32(obj) with _ -> 0

    member private this.groupKey(userId: string, noPaket: int64, noGroup: int) =
        userId + "|" + noPaket.ToString() + "|" + noGroup.ToString()

    member private this.finalizeExam(conn: SqlConnection, userId: string, noPaket: int64, idNoPeserta: int64, noPeserta: string) =
        use upd = new SqlCommand()
        upd.Connection <- conn
        upd.CommandType <- CommandType.Text
        upd.CommandText <-
            "UPDATE WISECON_PSIKOTEST.dbo.MS_PesertaDtl " +
            "SET TimeEdit = GETDATE(), UserEdit=@u, block=1, Url=NULL " +
            "WHERE UserId=@u AND NoPaket=@p; " +
            "IF COL_LENGTH('WISECON_PSIKOTEST.dbo.MS_PesertaDtl','bKirim') IS NOT NULL " +
            "BEGIN " +
            "  UPDATE WISECON_PSIKOTEST.dbo.MS_PesertaDtl " +
            "  SET bKirim = 1, UserEdit=@u, TimeEdit = GETDATE() " +
            "  WHERE UserId=@u AND NoPaket=@p; " +
            "END; " +
            "IF COL_LENGTH('WISECON_PSIKOTEST.dbo.MS_PesertaDtl','WaktuKirim') IS NOT NULL " +
            "BEGIN " +
            "  UPDATE WISECON_PSIKOTEST.dbo.MS_PesertaDtl " +
            "  SET WaktuKirim = ISNULL(WaktuKirim, GETDATE()), UserEdit=@u, TimeEdit = GETDATE() " +
            "  WHERE UserId=@u AND NoPaket=@p; " +
            "END;"
        upd.Parameters.AddWithValue("@u", userId) |> ignore
        upd.Parameters.AddWithValue("@p", noPaket) |> ignore
        upd.ExecuteNonQuery() |> ignore

        use delW = new SqlCommand("DELETE FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult WHERE UserId=@u OR NoPeserta=@n OR IdNoPeserta=@id", conn)
        delW.Parameters.AddWithValue("@u", userId) |> ignore
        delW.Parameters.AddWithValue("@n", noPeserta) |> ignore
        delW.Parameters.AddWithValue("@id", idNoPeserta) |> ignore
        delW.ExecuteNonQuery() |> ignore

        use hitung = new SqlCommand("dbo.SP_HitungNilaiUjian", conn)
        hitung.CommandType <- CommandType.StoredProcedure
        hitung.Parameters.AddWithValue("@UserId", userId) |> ignore
        hitung.ExecuteNonQuery() |> ignore

    member private this.loadQuestions(conn: SqlConnection, noPaket: int64) =
        use cmd = new SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <-
            "SELECT g.NoGroup, g.NamaGroup, g.MinimumJmlSoal, g.WaktuPengerjaan, ISNULL(g.bRandom,0) AS bRandom, ISNULL(g.IsPrioritas,0) AS IsPrioritas, " +
            "d.NoUrut, d.Judul, d.Deskripsi, " +
            "a.NoJawaban, a.Jawaban, CAST(a.NoJawabanBenar AS INT) AS PoinJawaban " +
            "FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup g " +
            "JOIN WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtl d ON d.NoPaket=g.NoPaket AND d.NoGroup=g.NoGroup " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtlJawaban a ON a.NoPaket=d.NoPaket AND a.NoGroup=d.NoGroup AND a.NoUrut=d.NoUrut " +
            "WHERE g.NoPaket=@p " +
            "ORDER BY ISNULL(g.IsPrioritas,0) DESC, g.NoGroup, d.NoUrut, a.NoJawaban"
        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
        use rdr = cmd.ExecuteReader()
        let groups = System.Collections.Generic.Dictionary<int, ExamGroup>()
        let qmaps = System.Collections.Generic.Dictionary<int, System.Collections.Generic.Dictionary<int, ExamQuestion>>()
        while rdr.Read() do
            let noGroup = if rdr.IsDBNull(0) then 0 else rdr.GetInt32(0)
            let namaGroup = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
            let minSoal = if rdr.IsDBNull(2) then 0 else rdr.GetInt32(2)
            let waktu = if rdr.IsDBNull(3) then 0 else rdr.GetInt32(3)
            let bRandom = if rdr.IsDBNull(4) then false else (try rdr.GetBoolean(4) with _ -> Convert.ToInt32(rdr.GetValue(4)) <> 0)
            let isPrioritas = if rdr.IsDBNull(5) then false else (try rdr.GetBoolean(5) with _ -> Convert.ToInt32(rdr.GetValue(5)) <> 0)
            let noUrut = if rdr.IsDBNull(6) then 0 else rdr.GetInt32(6)
            let judul = if rdr.IsDBNull(7) then "" else rdr.GetString(7)
            let deskripsi = if rdr.IsDBNull(8) then "" else rdr.GetString(8)
            let hasJawaban = not (rdr.IsDBNull(9))
            let noJawaban = if hasJawaban then (try int (rdr.GetByte(9)) with _ -> Convert.ToInt32(rdr.GetValue(9))) else -1
            let jawaban = if hasJawaban && not (rdr.IsDBNull(10)) then rdr.GetString(10) else ""
            let poin = if hasJawaban && not (rdr.IsDBNull(11)) then Convert.ToInt32(rdr.GetValue(11)) else 0

            if not (groups.ContainsKey(noGroup)) then
                groups.[noGroup] <- { noGroup = noGroup; namaGroup = namaGroup; minSoal = minSoal; waktu = waktu; random = bRandom; isPrioritas = isPrioritas; questions = ResizeArray<ExamQuestion>() }
                qmaps.[noGroup] <- System.Collections.Generic.Dictionary<int, ExamQuestion>()

            let qmap = qmaps.[noGroup]
            if not (qmap.ContainsKey(noUrut)) then
                let q = { noUrut = noUrut; judul = judul; deskripsi = deskripsi; answers = ResizeArray<ExamAnswer>() }
                qmap.[noUrut] <- q
                groups.[noGroup].questions.Add(q)

            let qObj = qmap.[noUrut]
            if hasJawaban then
                qObj.answers.Add({ noJawaban = noJawaban; jawaban = jawaban; poin = poin })

        groups.Values |> Seq.toArray

    [<AllowAnonymous>]
    [<HttpPost>]
    [<Route("Exam/Finish")>]
    member this.Finish([<FromForm>] req: ExamFinishRequest) : IActionResult =
        let encKey = getEncryptionKey ()
        if String.IsNullOrWhiteSpace(encKey) then
            this.BadRequest("Konfigurasi EncryptionKey belum diset") :> IActionResult
        else
            let tokenEnc = norm req.Token
            let tokenEncFixed = Uri.UnescapeDataString(tokenEnc).Replace(" ", "+")
            let plain = decrypt64 tokenEncFixed encKey
            match tryParseToken plain with
            | None -> this.BadRequest("Token tidak valid") :> IActionResult
            | Some(userId, noPaket, waktuFromToken, _, _) ->
                let conn = db :?> SqlConnection
                conn.Open()
                try
                    match this.getAssignment(conn, userId, noPaket) with
                    | None -> this.BadRequest("Data ujian tidak ditemukan") :> IActionResult
                    | Some(idNoPeserta, noPeserta, _, _, waktuTest, startTest, timeEdit, isBlocked, _, _, _) ->
                        if isBlocked then
                            this.BadRequest("Akses ujian diblokir") :> IActionResult
                        else
                            let now = DateTime.Now
                            let plannedTime =
                                if waktuTest.HasValue then waktuTest.Value
                                else waktuFromToken
                            let toleransi = this.getToleranceMinutes(conn, noPaket)
                            let startWindow = plannedTime.AddMinutes(-float toleransi)
                            let endWindow = plannedTime.AddMinutes(float toleransi)
                            if not startTest.HasValue && (now < startWindow || now > endWindow) then
                                this.BadRequest("Waktu ujian tidak valid") :> IActionResult
                            else
                                try
                                    this.finalizeExam(conn, userId, noPaket, idNoPeserta, noPeserta)
                                    let keyPrefix = userId + "|" + noPaket.ToString() + "|"
                                    for kv in groupStartMap do
                                        if kv.Key.StartsWith(keyPrefix, StringComparison.Ordinal) then
                                            groupStartMap.TryRemove(kv.Key) |> ignore
                                    this.Ok(box {| ok = true |}) :> IActionResult
                                with ex ->
                                    this.BadRequest(ex.Message) :> IActionResult
                finally
                    conn.Close()

    [<AllowAnonymous>]
    [<HttpPost>]
    [<Route("Exam/Answer")>]
    member this.SubmitAnswer([<FromBody>] req: ExamSubmitAnswerRequest) : IActionResult =
        let encKey = getEncryptionKey ()
        if String.IsNullOrWhiteSpace(encKey) then
            this.BadRequest("Konfigurasi EncryptionKey belum diset") :> IActionResult
        else
            let tokenEnc = norm req.Token
            let tokenEncFixed = Uri.UnescapeDataString(tokenEnc).Replace(" ", "+")
            let plain = decrypt64 tokenEncFixed encKey
            match tryParseToken plain with
            | None -> this.BadRequest("Token tidak valid") :> IActionResult
            | Some(userId, noPaket, waktuFromToken, _, _) ->
                let conn = db :?> SqlConnection
                conn.Open()
                try
                    match this.getAssignment(conn, userId, noPaket) with
                    | None -> this.BadRequest("Data ujian tidak ditemukan") :> IActionResult
                    | Some(idNoPeserta, noPeserta, _, _, waktuTest, startTest, timeEdit, isBlocked, _, _, _) ->
                        if isBlocked then this.BadRequest("Akses ujian diblokir") :> IActionResult
                        else
                            let now = DateTime.Now
                            if this.hasFinalResult(conn, userId) then
                                this.BadRequest("Ujian sudah selesai") :> IActionResult
                            else
                                let plannedTime =
                                    if waktuTest.HasValue then waktuTest.Value
                                    else waktuFromToken
                                let toleransi = this.getToleranceMinutes(conn, noPaket)
                                let startWindow = plannedTime.AddMinutes(-float toleransi)
                                let endWindow = plannedTime.AddMinutes(float toleransi)
                                if not startTest.HasValue && (now < startWindow || now > endWindow) then
                                    this.BadRequest("Waktu ujian tidak valid") :> IActionResult
                                else
                                    let duration = this.getGroupDurationMinutes(conn, noPaket, req.NoGroup)
                                    if duration > 0 then
                                        let key = this.groupKey(userId, noPaket, req.NoGroup)
                                        let startAt =
                                            match groupStartMap.TryGetValue(key) with
                                            | true, v -> v
                                            | _ ->
                                                let v = DateTime.Now
                                                groupStartMap.[key] <- v
                                                v
                                        let endAt = startAt.AddMinutes(float duration)
                                        if now > endAt.AddSeconds(5.0) then
                                            this.StatusCode(410, "Waktu group sudah habis") :> IActionResult
                                        else
                                            use cmdTipe = new SqlCommand()
                                            cmdTipe.Connection <- conn
                                            cmdTipe.CommandType <- CommandType.Text
                                            cmdTipe.CommandText <- "SELECT ISNULL(Tipe,'PG') FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup WHERE NoPaket=@p AND NoGroup=@g"
                                            cmdTipe.Parameters.AddWithValue("@p", noPaket) |> ignore
                                            cmdTipe.Parameters.AddWithValue("@g", req.NoGroup) |> ignore
                                            let tipeObj = cmdTipe.ExecuteScalar()
                                            let tipe = if isNull tipeObj then "PG" else Convert.ToString(tipeObj)
                                            if String.Equals(tipe, "PG", StringComparison.OrdinalIgnoreCase) then
                                                use chk = new SqlCommand()
                                                chk.Connection <- conn
                                                chk.CommandType <- CommandType.Text
                                                chk.CommandText <- "SELECT CASE WHEN EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtlJawaban WHERE NoPaket=@p AND NoGroup=@g AND NoUrut=@u AND NoJawaban=@j) THEN 1 ELSE 0 END"
                                                chk.Parameters.AddWithValue("@p", noPaket) |> ignore
                                                chk.Parameters.AddWithValue("@g", req.NoGroup) |> ignore
                                                chk.Parameters.AddWithValue("@u", req.NoUrut) |> ignore
                                                chk.Parameters.AddWithValue("@j", req.JawabanDiPilih) |> ignore
                                                let okObj = chk.ExecuteScalar()
                                                let ok = try Convert.ToInt32(okObj) = 1 with _ -> false
                                                if not ok then
                                                    this.BadRequest("Pilihan jawaban tidak valid") :> IActionResult
                                                else
                                                    try
                                                        use sp = new SqlCommand("dbo.SP_SubmitJawaban", conn)
                                                        sp.CommandType <- CommandType.StoredProcedure
                                                        sp.Parameters.AddWithValue("@NoPeserta", idNoPeserta) |> ignore
                                                        sp.Parameters.AddWithValue("@User", userId) |> ignore
                                                        sp.Parameters.AddWithValue("@Tipe", tipe) |> ignore
                                                        sp.Parameters.AddWithValue("@NoPaket", Convert.ToInt32(noPaket)) |> ignore
                                                        sp.Parameters.AddWithValue("@NoGroup", req.NoGroup) |> ignore
                                                        sp.Parameters.AddWithValue("@NoUrut", req.NoUrut) |> ignore
                                                        sp.Parameters.AddWithValue("@JawabanDiPilih", req.JawabanDiPilih) |> ignore
                                                        sp.Parameters.AddWithValue("@JmlSalah", 0uy) |> ignore
                                                        sp.ExecuteNonQuery() |> ignore
                                                        this.Ok(box {| ok = true |}) :> IActionResult
                                                    with ex ->
                                                        this.BadRequest(ex.Message) :> IActionResult
                                            else
                                                try
                                                    use sp = new SqlCommand("dbo.SP_SubmitJawaban", conn)
                                                    sp.CommandType <- CommandType.StoredProcedure
                                                    sp.Parameters.AddWithValue("@NoPeserta", idNoPeserta) |> ignore
                                                    sp.Parameters.AddWithValue("@User", userId) |> ignore
                                                    sp.Parameters.AddWithValue("@Tipe", tipe) |> ignore
                                                    sp.Parameters.AddWithValue("@NoPaket", Convert.ToInt32(noPaket)) |> ignore
                                                    sp.Parameters.AddWithValue("@NoGroup", req.NoGroup) |> ignore
                                                    sp.Parameters.AddWithValue("@NoUrut", req.NoUrut) |> ignore
                                                    sp.Parameters.AddWithValue("@JawabanDiPilih", req.JawabanDiPilih) |> ignore
                                                    sp.Parameters.AddWithValue("@JmlSalah", 0uy) |> ignore
                                                    sp.ExecuteNonQuery() |> ignore
                                                    this.Ok(box {| ok = true |}) :> IActionResult
                                                with ex ->
                                                    this.BadRequest(ex.Message) :> IActionResult
                                    else
                                        use cmdTipe = new SqlCommand()
                                        cmdTipe.Connection <- conn
                                        cmdTipe.CommandType <- CommandType.Text
                                        cmdTipe.CommandText <- "SELECT ISNULL(Tipe,'PG') FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup WHERE NoPaket=@p AND NoGroup=@g"
                                        cmdTipe.Parameters.AddWithValue("@p", noPaket) |> ignore
                                        cmdTipe.Parameters.AddWithValue("@g", req.NoGroup) |> ignore
                                        let tipeObj = cmdTipe.ExecuteScalar()
                                        let tipe = if isNull tipeObj then "PG" else Convert.ToString(tipeObj)
                                        if String.Equals(tipe, "PG", StringComparison.OrdinalIgnoreCase) then
                                            use chk = new SqlCommand()
                                            chk.Connection <- conn
                                            chk.CommandType <- CommandType.Text
                                            chk.CommandText <- "SELECT CASE WHEN EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroupDtlJawaban WHERE NoPaket=@p AND NoGroup=@g AND NoUrut=@u AND NoJawaban=@j) THEN 1 ELSE 0 END"
                                            chk.Parameters.AddWithValue("@p", noPaket) |> ignore
                                            chk.Parameters.AddWithValue("@g", req.NoGroup) |> ignore
                                            chk.Parameters.AddWithValue("@u", req.NoUrut) |> ignore
                                            chk.Parameters.AddWithValue("@j", req.JawabanDiPilih) |> ignore
                                            let okObj = chk.ExecuteScalar()
                                            let ok = try Convert.ToInt32(okObj) = 1 with _ -> false
                                            if not ok then
                                                this.BadRequest("Pilihan jawaban tidak valid") :> IActionResult
                                            else
                                                try
                                                    use sp = new SqlCommand("dbo.SP_SubmitJawaban", conn)
                                                    sp.CommandType <- CommandType.StoredProcedure
                                                    sp.Parameters.AddWithValue("@NoPeserta", idNoPeserta) |> ignore
                                                    sp.Parameters.AddWithValue("@User", userId) |> ignore
                                                    sp.Parameters.AddWithValue("@Tipe", tipe) |> ignore
                                                    sp.Parameters.AddWithValue("@NoPaket", Convert.ToInt32(noPaket)) |> ignore
                                                    sp.Parameters.AddWithValue("@NoGroup", req.NoGroup) |> ignore
                                                    sp.Parameters.AddWithValue("@NoUrut", req.NoUrut) |> ignore
                                                    sp.Parameters.AddWithValue("@JawabanDiPilih", req.JawabanDiPilih) |> ignore
                                                    sp.Parameters.AddWithValue("@JmlSalah", 0uy) |> ignore
                                                    sp.ExecuteNonQuery() |> ignore
                                                    this.Ok(box {| ok = true |}) :> IActionResult
                                                with ex ->
                                                    this.BadRequest(ex.Message) :> IActionResult
                                        else
                                            try
                                                use sp = new SqlCommand("dbo.SP_SubmitJawaban", conn)
                                                sp.CommandType <- CommandType.StoredProcedure
                                                sp.Parameters.AddWithValue("@NoPeserta", idNoPeserta) |> ignore
                                                sp.Parameters.AddWithValue("@User", userId) |> ignore
                                                sp.Parameters.AddWithValue("@Tipe", tipe) |> ignore
                                                sp.Parameters.AddWithValue("@NoPaket", Convert.ToInt32(noPaket)) |> ignore
                                                sp.Parameters.AddWithValue("@NoGroup", req.NoGroup) |> ignore
                                                sp.Parameters.AddWithValue("@NoUrut", req.NoUrut) |> ignore
                                                sp.Parameters.AddWithValue("@JawabanDiPilih", req.JawabanDiPilih) |> ignore
                                                sp.Parameters.AddWithValue("@JmlSalah", 0uy) |> ignore
                                                sp.ExecuteNonQuery() |> ignore
                                                this.Ok(box {| ok = true |}) :> IActionResult
                                            with ex ->
                                                this.BadRequest(ex.Message) :> IActionResult
                finally
                    conn.Close()

    [<AllowAnonymous>]
    [<HttpPost>]
    [<Route("Exam/GroupStart")>]
    member this.GroupStart([<FromBody>] req: ExamGroupStartRequest) : IActionResult =
        let encKey = getEncryptionKey ()
        if String.IsNullOrWhiteSpace(encKey) then
            this.BadRequest("Konfigurasi EncryptionKey belum diset") :> IActionResult
        else
            let tokenEnc = norm req.Token
            let tokenEncFixed = Uri.UnescapeDataString(tokenEnc).Replace(" ", "+")
            let plain = decrypt64 tokenEncFixed encKey
            match tryParseToken plain with
            | None -> this.BadRequest("Token tidak valid") :> IActionResult
            | Some(userId, noPaket, _, _, _) ->
                let conn = db :?> SqlConnection
                conn.Open()
                try
                    match this.getAssignment(conn, userId, noPaket) with
                    | None -> this.BadRequest("Data ujian tidak ditemukan") :> IActionResult
                    | Some(_, _, _, _, _, _, _, isBlocked, _, _, _) ->
                        if isBlocked then this.BadRequest("Akses ujian diblokir") :> IActionResult
                        else
                            let key = this.groupKey(userId, noPaket, req.NoGroup)
                            groupStartMap.[key] <- DateTime.Now
                            this.Ok(box {| ok = true |}) :> IActionResult
                finally
                    conn.Close()

    [<AllowAnonymous>]
    [<HttpGet>]
    [<Route("Exam")>]
    member this.Index([<FromQuery>] Token: string) : IActionResult =
        this.ViewData.["Token"] <- Token
        this.View()

    [<AllowAnonymous>]
    [<HttpPost>]
    [<Route("Exam/Login")>]
    member this.Login([<FromForm>] req: ExamLoginRequest) : IActionResult =
        let encKey = getEncryptionKey ()
        if String.IsNullOrWhiteSpace(encKey) then
            this.ViewData.["ErrorMessage"] <- "Konfigurasi EncryptionKey belum diset"
            this.View("Index") :> IActionResult
        else
            let tokenEnc = norm req.Token
            let tokenEncFixed = Uri.UnescapeDataString(tokenEnc).Replace(" ", "+")
            let plain = decrypt64 tokenEncFixed encKey
            match tryParseToken plain with
            | None ->
                this.ViewData.["ErrorMessage"] <- "Token tidak valid"
                this.ViewData.["Token"] <- req.Token
                this.View("Index") :> IActionResult
            | Some(userId, noPaket, waktuFromToken, noKtpToken, emailToken) ->
                let conn = db :?> SqlConnection
                conn.Open()
                try
                    match this.getAssignment(conn, userId, noPaket) with
                    | None ->
                        this.ViewData.["ErrorMessage"] <- "Data ujian tidak ditemukan"
                        this.ViewData.["Token"] <- req.Token
                        this.View("Index") :> IActionResult
                    | Some(idNoPeserta, noPeserta, _, _, waktuTest, startTest, timeEdit, isBlocked, noKtpDb, emailDb, namaPaket) ->
                        if isBlocked then
                            this.ViewData.["ErrorMessage"] <- "Akses ujian diblokir"
                            this.ViewData.["Token"] <- req.Token
                            this.View("Index") :> IActionResult
                        else
                            let noKtpIn = norm req.NoKTP
                            let emailIn = norm req.Email
                            if String.IsNullOrWhiteSpace(noKtpIn) || String.IsNullOrWhiteSpace(emailIn) then
                                this.ViewData.["ErrorMessage"] <- "No KTP dan Email wajib diisi"
                                this.ViewData.["Token"] <- req.Token
                                this.View("Index") :> IActionResult
                            elif (not (String.IsNullOrWhiteSpace(noKtpToken)) && not (eqi noKtpIn noKtpToken)) || (not (String.IsNullOrWhiteSpace(emailToken)) && not (eqi emailIn emailToken)) then
                                this.ViewData.["ErrorMessage"] <- "No KTP atau Email tidak sesuai"
                                this.ViewData.["Token"] <- req.Token
                                this.View("Index") :> IActionResult
                            elif not (eqi noKtpIn noKtpDb) || not (eqi emailIn emailDb) then
                                this.ViewData.["ErrorMessage"] <- "No KTP atau Email tidak sesuai"
                                this.ViewData.["Token"] <- req.Token
                                this.View("Index") :> IActionResult
                            else
                                if this.hasFinalResult(conn, userId) then
                                    this.ViewData.["ErrorMessage"] <- "Ujian sudah selesai"
                                    this.ViewData.["Token"] <- req.Token
                                    this.View("Index") :> IActionResult
                                else
                                    let plannedTime =
                                        if waktuTest.HasValue then waktuTest.Value
                                        else waktuFromToken
                                    let toleransi = this.getToleranceMinutes(conn, noPaket)
                                    let startWindow = plannedTime.AddMinutes(-float toleransi)
                                    let endWindow = plannedTime.AddMinutes(float toleransi)
                                    let now = DateTime.Now
                                    if not startTest.HasValue && (now < startWindow || now > endWindow) then
                                        this.ViewData.["ErrorMessage"] <- "Waktu ujian tidak valid"
                                        this.ViewData.["Token"] <- req.Token
                                        this.View("Index") :> IActionResult
                                    else
                                        use cmd = new SqlCommand()
                                        cmd.Connection <- conn
                                        cmd.CommandType <- CommandType.Text
                                        cmd.CommandText <- "UPDATE WISECON_PSIKOTEST.dbo.MS_PesertaDtl SET StartTest = ISNULL(StartTest, GETDATE()), UserEdit=@u WHERE UserId=@u AND NoPaket=@p"
                                        cmd.Parameters.AddWithValue("@u", userId) |> ignore
                                        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
                                        let _ = cmd.ExecuteNonQuery()
                                        let q = this.loadQuestions(conn, noPaket)
                                        use cmdStart = new SqlCommand("SELECT TOP 1 StartTest FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl WHERE UserId=@u AND NoPaket=@p ORDER BY TimeInput DESC", conn)
                                        cmdStart.Parameters.AddWithValue("@u", userId) |> ignore
                                        cmdStart.Parameters.AddWithValue("@p", noPaket) |> ignore
                                        let startObj = cmdStart.ExecuteScalar()
                                        let startAt2 =
                                            if isNull startObj then now
                                            else
                                                try Convert.ToDateTime(startObj)
                                                with _ -> now
                                        this.ViewData.["Token"] <- req.Token
                                        this.ViewData.["UserId"] <- userId
                                        this.ViewData.["NoPaket"] <- noPaket
                                        this.ViewData.["NoPeserta"] <- noPeserta
                                        this.ViewData.["NamaPaket"] <- namaPaket
                                        this.ViewData.["StartAtIso"] <- startAt2.ToString("yyyy-MM-ddTHH:mm:ss")
                                        this.ViewData.["EndAtIso"] <- startAt2.ToString("yyyy-MM-ddTHH:mm:ss")
                                        this.ViewData.["QuestionsJson"] <- System.Text.Json.JsonSerializer.Serialize(q)
                                        this.View("Test") :> IActionResult
                finally
                    conn.Close()
