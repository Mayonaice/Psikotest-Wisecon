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

[<CLIMutable>]
type ExamGroupCompleteRequest = { Token: string; NoGroup: int }

type ExamAnswer = { noJawaban: int; jawaban: string; poin: int; tipeMedia: string; urlMedia: string; textMedia: string }
type ExamQuestion = { noUrut: int; judul: string; deskripsi: string; tipeMedia: string; urlMedia: string; answers: ResizeArray<ExamAnswer> }
type ExamGroup = { noGroup: int; namaGroup: string; minSoal: int; waktu: int; random: bool; isPrioritas: bool; noPetunjuk: int; petunjuk: string; questions: ResizeArray<ExamQuestion> }

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

    member private this.getLastIncompleteGroup(conn: SqlConnection, userId: string, noPaket: int64, attemptTime: DateTime) =
        use cmd = new SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- 
            "SELECT TOP 1 NoGroup FROM WISECON_PSIKOTEST.dbo.TR_PesertaGroupProgress " +
            "WHERE UserId=@u AND NoPaket=@p AND IsCompleted=0 " +
            "AND CONVERT(VARCHAR(19), AttemptTime, 120) = CONVERT(VARCHAR(19), @at, 120) " +
            "ORDER BY TimeInput ASC"
        cmd.Parameters.AddWithValue("@u", userId) |> ignore
        cmd.Parameters.AddWithValue("@p", noPaket) |> ignore
        cmd.Parameters.AddWithValue("@at", attemptTime) |> ignore
        let obj = cmd.ExecuteScalar()
        if isNull obj then -1 else try Convert.ToInt32(obj) with _ -> -1

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

        use insHist = new SqlCommand()
        insHist.Connection <- conn
        insHist.CommandType <- CommandType.Text
        insHist.CommandText <-
            "IF OBJECT_ID('WISECON_PSIKOTEST.dbo.TR_PsikotestResultHistory','U') IS NOT NULL " +
            "BEGIN " +
            "  INSERT INTO WISECON_PSIKOTEST.dbo.TR_PsikotestResultHistory " +
            "  (IdNoPeserta, NoPeserta, UserId, NoPaket, UndangPsikotestKe, AttemptTime, GroupSoal, NilaiStandard, NilaiGroupResult, UserInput, TimeInput, UserEdit, TimeEdit, ArchivedAt) " +
            "  SELECT R.IdNoPeserta, R.NoPeserta, R.UserId, D.NoPaket, ISNULL(D.UndangPsikotestKe,0), D.TimeInput, " +
            "         R.GroupSoal, R.NilaiStandard, R.NilaiGroupResult, R.UserInput, R.TimeInput, R.UserEdit, R.TimeEdit, GETDATE() " +
            "  FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult R " +
            "  OUTER APPLY (SELECT TOP 1 NoPaket, UndangPsikotestKe, TimeInput FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl WHERE (UserId=R.UserId AND ISNULL(R.UserId,'')<>'') OR (ISNULL(R.UserId,'')='' AND NoPeserta=R.NoPeserta) ORDER BY TimeInput DESC) D " +
            "  WHERE (R.UserId=@u OR R.NoPeserta=@n OR R.IdNoPeserta=@id) " +
            "    AND NOT EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResultHistory H WHERE H.IdNoPeserta=R.IdNoPeserta AND H.GroupSoal=R.GroupSoal AND H.AttemptTime=COALESCE(D.TimeInput, R.TimeInput)); " +
            "END;"
        insHist.Parameters.AddWithValue("@u", userId) |> ignore
        insHist.Parameters.AddWithValue("@p", noPaket) |> ignore
        insHist.Parameters.AddWithValue("@n", noPeserta) |> ignore
        insHist.Parameters.AddWithValue("@id", idNoPeserta) |> ignore
        insHist.ExecuteNonQuery() |> ignore

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
            "ISNULL(g.NoPetunjuk,0) AS NoPetunjuk, ISNULL(p.Keterangan,'') AS Keterangan, " +
            "d.NoUrut, d.Judul, d.Deskripsi, ISNULL(d.UrlMedia,'') AS DUrlMedia, ISNULL(d.TipeMedia,'NOMEDIA') AS DTipeMedia, " +
            "a.NoJawaban, a.Jawaban, CAST(a.NoJawabanBenar AS INT) AS PoinJawaban, ISNULL(a.UrlMedia,'') AS AUrlMedia, ISNULL(a.TipeMedia,'NOMEDIA') AS ATipeMedia, ISNULL(a.TextMedia,'') AS ATextMedia " +
            "FROM WISECON_PSIKOTEST.dbo.MS_PaketSoalGroup g " +
            "LEFT JOIN WISECON_PSIKOTEST.dbo.MS_Petunjuk p ON p.SeqNo=g.NoPetunjuk " +
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
            let noPetunjuk = if rdr.IsDBNull(6) then 0 else Convert.ToInt32(rdr.GetValue(6))
            let petunjuk = if rdr.IsDBNull(7) then "" else rdr.GetString(7)
            let noUrut = if rdr.IsDBNull(8) then 0 else rdr.GetInt32(8)
            let judul = if rdr.IsDBNull(9) then "" else rdr.GetString(9)
            let deskripsi = if rdr.IsDBNull(10) then "" else rdr.GetString(10)
            let qUrlMedia = if rdr.IsDBNull(11) then "" else rdr.GetString(11)
            let qTipeMedia = if rdr.IsDBNull(12) then "NOMEDIA" else rdr.GetString(12)
            let hasJawaban = not (rdr.IsDBNull(13))
            let noJawaban = if hasJawaban then (try int (rdr.GetByte(13)) with _ -> Convert.ToInt32(rdr.GetValue(13))) else -1
            let jawaban = if hasJawaban && not (rdr.IsDBNull(14)) then rdr.GetString(14) else ""
            let poin = if hasJawaban && not (rdr.IsDBNull(15)) then Convert.ToInt32(rdr.GetValue(15)) else 0
            let aUrlMedia = if hasJawaban && not (rdr.IsDBNull(16)) then rdr.GetString(16) else ""
            let aTipeMedia = if hasJawaban && not (rdr.IsDBNull(17)) then rdr.GetString(17) else "NOMEDIA"
            let aTextMedia = if hasJawaban && not (rdr.IsDBNull(18)) then rdr.GetString(18) else ""

            if not (groups.ContainsKey(noGroup)) then
                groups.[noGroup] <- { noGroup = noGroup; namaGroup = namaGroup; minSoal = minSoal; waktu = waktu; random = bRandom; isPrioritas = isPrioritas; noPetunjuk = noPetunjuk; petunjuk = petunjuk; questions = ResizeArray<ExamQuestion>() }
                qmaps.[noGroup] <- System.Collections.Generic.Dictionary<int, ExamQuestion>()

            let qmap = qmaps.[noGroup]
            if not (qmap.ContainsKey(noUrut)) then
                let q = { noUrut = noUrut; judul = judul; deskripsi = deskripsi; tipeMedia = qTipeMedia; urlMedia = qUrlMedia; answers = ResizeArray<ExamAnswer>() }
                qmap.[noUrut] <- q
                groups.[noGroup].questions.Add(q)

            let qObj = qmap.[noUrut]
            if hasJawaban then
                qObj.answers.Add({ noJawaban = noJawaban; jawaban = jawaban; poin = poin; tipeMedia = aTipeMedia; urlMedia = aUrlMedia; textMedia = aTextMedia })

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
                            
                            // Get attempt time from MS_PesertaDtl
                            use cmdAttempt = new SqlCommand()
                            cmdAttempt.Connection <- conn
                            cmdAttempt.CommandType <- CommandType.Text
                            cmdAttempt.CommandText <- 
                                "SELECT TOP 1 TimeInput FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl " +
                                "WHERE UserId=@u AND NoPaket=@p ORDER BY TimeInput DESC"
                            cmdAttempt.Parameters.AddWithValue("@u", userId) |> ignore
                            cmdAttempt.Parameters.AddWithValue("@p", noPaket) |> ignore
                            let attemptObj = cmdAttempt.ExecuteScalar()
                            let attemptTime = 
                                if isNull attemptObj then DateTime.Now
                                else try Convert.ToDateTime(attemptObj) with _ -> DateTime.Now
                            
                            // Mark group as started in progress table
                            use cmdCheck = new SqlCommand()
                            cmdCheck.Connection <- conn
                            cmdCheck.CommandType <- CommandType.Text
                            cmdCheck.CommandText <- 
                                "SELECT COUNT(*) FROM WISECON_PSIKOTEST.dbo.TR_PesertaGroupProgress " +
                                "WHERE UserId=@u AND NoPaket=@p AND NoGroup=@g " +
                                "AND CONVERT(VARCHAR(19), AttemptTime, 120) = CONVERT(VARCHAR(19), @at, 120)"
                            cmdCheck.Parameters.AddWithValue("@u", userId) |> ignore
                            cmdCheck.Parameters.AddWithValue("@p", noPaket) |> ignore
                            cmdCheck.Parameters.AddWithValue("@g", req.NoGroup) |> ignore
                            cmdCheck.Parameters.AddWithValue("@at", attemptTime) |> ignore
                            let count = Convert.ToInt32(cmdCheck.ExecuteScalar())
                            
                            if count = 0 then
                                use cmdInsert = new SqlCommand()
                                cmdInsert.Connection <- conn
                                cmdInsert.CommandType <- CommandType.Text
                                cmdInsert.CommandText <-
                                    "INSERT INTO WISECON_PSIKOTEST.dbo.TR_PesertaGroupProgress " +
                                    "(UserId, NoPaket, NoGroup, IsCompleted, AttemptTime, UserInput, TimeInput) " +
                                    "VALUES (@u, @p, @g, 0, @at, @u, GETDATE())"
                                cmdInsert.Parameters.AddWithValue("@u", userId) |> ignore
                                cmdInsert.Parameters.AddWithValue("@p", noPaket) |> ignore
                                cmdInsert.Parameters.AddWithValue("@g", req.NoGroup) |> ignore
                                cmdInsert.Parameters.AddWithValue("@at", attemptTime) |> ignore
                                cmdInsert.ExecuteNonQuery() |> ignore
                            
                            this.Ok(box {| ok = true |}) :> IActionResult
                finally
                    conn.Close()

    [<AllowAnonymous>]
    [<HttpPost>]
    [<Route("Exam/GroupComplete")>]
    member this.GroupComplete([<FromBody>] req: ExamGroupCompleteRequest) : IActionResult =
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
                            // Get attempt time from MS_PesertaDtl
                            use cmdAttempt = new SqlCommand()
                            cmdAttempt.Connection <- conn
                            cmdAttempt.CommandType <- CommandType.Text
                            cmdAttempt.CommandText <- 
                                "SELECT TOP 1 TimeInput FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl " +
                                "WHERE UserId=@u AND NoPaket=@p ORDER BY TimeInput DESC"
                            cmdAttempt.Parameters.AddWithValue("@u", userId) |> ignore
                            cmdAttempt.Parameters.AddWithValue("@p", noPaket) |> ignore
                            let attemptObj = cmdAttempt.ExecuteScalar()
                            let attemptTime = 
                                if isNull attemptObj then DateTime.Now
                                else try Convert.ToDateTime(attemptObj) with _ -> DateTime.Now
                            
                            // Mark group as completed
                            use cmdUpdate = new SqlCommand()
                            cmdUpdate.Connection <- conn
                            cmdUpdate.CommandType <- CommandType.Text
                            cmdUpdate.CommandText <-
                                "UPDATE WISECON_PSIKOTEST.dbo.TR_PesertaGroupProgress " +
                                "SET IsCompleted=1, CompletedAt=GETDATE() " +
                                "WHERE UserId=@u AND NoPaket=@p AND NoGroup=@g " +
                                "AND CONVERT(VARCHAR(19), AttemptTime, 120) = CONVERT(VARCHAR(19), @at, 120)"
                            cmdUpdate.Parameters.AddWithValue("@u", userId) |> ignore
                            cmdUpdate.Parameters.AddWithValue("@p", noPaket) |> ignore
                            cmdUpdate.Parameters.AddWithValue("@g", req.NoGroup) |> ignore
                            cmdUpdate.Parameters.AddWithValue("@at", attemptTime) |> ignore
                            cmdUpdate.ExecuteNonQuery() |> ignore
                            
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
                                        
                                        use cmdStart = new SqlCommand("SELECT TOP 1 StartTest, TimeInput FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl WHERE UserId=@u AND NoPaket=@p ORDER BY TimeInput DESC", conn)
                                        cmdStart.Parameters.AddWithValue("@u", userId) |> ignore
                                        cmdStart.Parameters.AddWithValue("@p", noPaket) |> ignore
                                        use rdr = cmdStart.ExecuteReader()
                                        let (startAt2, attemptTime) =
                                            if rdr.Read() then
                                                let st = if rdr.IsDBNull(0) then now else rdr.GetDateTime(0)
                                                let at = if rdr.IsDBNull(1) then now else rdr.GetDateTime(1)
                                                (st, at)
                                            else
                                                (now, now)
                                        rdr.Close()
                                        
                                        // Get last incomplete group for this attempt
                                        let lastIncompleteGroup = this.getLastIncompleteGroup(conn, userId, noPaket, attemptTime)
                                        
                                        this.ViewData.["Token"] <- req.Token
                                        this.ViewData.["UserId"] <- userId
                                        this.ViewData.["NoPaket"] <- noPaket
                                        this.ViewData.["NoPeserta"] <- noPeserta
                                        this.ViewData.["NamaPaket"] <- namaPaket
                                        this.ViewData.["StartAtIso"] <- startAt2.ToString("yyyy-MM-ddTHH:mm:ss")
                                        this.ViewData.["EndAtIso"] <- startAt2.ToString("yyyy-MM-ddTHH:mm:ss")
                                        this.ViewData.["QuestionsJson"] <- System.Text.Json.JsonSerializer.Serialize(q)
                                        this.ViewData.["LastIncompleteGroup"] <- lastIncompleteGroup
                                        this.View("Test") :> IActionResult
                finally
                    conn.Close()

    [<AllowAnonymous>]
    [<HttpGet>]
    [<Route("Exam/GetPhotoInterval")>]
    member this.GetPhotoInterval() : IActionResult =
        let conn = db :?> SqlConnection
        conn.Open()
        try
            use cmd = new SqlCommand()
            cmd.Connection <- conn
            cmd.CommandType <- CommandType.Text
            cmd.CommandText <- "SELECT TOP 1 Value FROM WISECON_PSIKOTEST.dbo.SYS_Parameter WHERE Name = 'Interval_Waktu_Foto' AND Status_Parameter = '1'"
            let obj = cmd.ExecuteScalar()
            if isNull obj then
                this.Ok(box {| interval = 0 |}) :> IActionResult
            else
                let intervalStr = obj.ToString()
                let mutable intervalMin = 0
                if Int32.TryParse(intervalStr, &intervalMin) then
                    this.Ok(box {| interval = intervalMin |}) :> IActionResult
                else
                    this.Ok(box {| interval = 0 |}) :> IActionResult
        finally
            conn.Close()

    [<AllowAnonymous>]
    [<HttpPost>]
    [<Route("Exam/SavePhoto")>]
    member this.SavePhoto() : IActionResult =
        try
            let form = this.Request.Form
            let token = if form.ContainsKey("token") then form.["token"].ToString() else ""
            let photoData = if form.ContainsKey("photo") then form.["photo"].ToString() else ""
            
            if String.IsNullOrWhiteSpace(token) || String.IsNullOrWhiteSpace(photoData) then
                this.BadRequest("Token atau photo data kosong") :> IActionResult
            else
                let encKey = getEncryptionKey ()
                if String.IsNullOrWhiteSpace(encKey) then
                    this.BadRequest("Konfigurasi EncryptionKey belum diset") :> IActionResult
                else
                    let tokenEncFixed = Uri.UnescapeDataString(token).Replace(" ", "+")
                    let plain = decrypt64 tokenEncFixed encKey
                    match tryParseToken plain with
                    | None -> this.BadRequest("Token tidak valid") :> IActionResult
                    | Some(userId, noPaket, _, _, _) ->
                        let conn = db :?> SqlConnection
                        conn.Open()
                        try
                            // Get IdNoPeserta and NoPeserta
                            match this.getAssignment(conn, userId, noPaket) with
                            | None -> this.BadRequest("Data ujian tidak ditemukan") :> IActionResult
                            | Some(idNoPeserta, noPeserta, _, _, _, _, _, _, _, _, _) ->
                                // Remove data:image/png;base64, prefix
                                let base64Data = 
                                    if photoData.Contains(",") then
                                        photoData.Substring(photoData.IndexOf(",") + 1)
                                    else photoData
                                
                                let imageBytes = Convert.FromBase64String(base64Data)
                                let basePath = @"D:\dct_docs\WISECON_PSIKOTEST\Test\WajahPeserta"
                                
                                // Create directory if not exists
                                if not (System.IO.Directory.Exists(basePath)) then
                                    System.IO.Directory.CreateDirectory(basePath) |> ignore
                                
                                // Generate filename: UserId_NoPaket_Timestamp.jpg
                                let timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss_fff")
                                let filename = sprintf "%s_%d_%s.jpg" userId noPaket timestamp
                                let fullPath = System.IO.Path.Combine(basePath, filename)
                                
                                // Save file
                                System.IO.File.WriteAllBytes(fullPath, imageBytes)
                                
                                // Create table if not exists and insert record
                                use cmdCreateTable = new SqlCommand()
                                cmdCreateTable.Connection <- conn
                                cmdCreateTable.CommandType <- CommandType.Text
                                cmdCreateTable.CommandText <-
                                    "IF OBJECT_ID('WISECON_PSIKOTEST.dbo.TR_PesertaFoto','U') IS NULL " +
                                    "BEGIN " +
                                    "  CREATE TABLE WISECON_PSIKOTEST.dbo.TR_PesertaFoto ( " +
                                    "    ID BIGINT IDENTITY(1,1) PRIMARY KEY, " +
                                    "    IdNoPeserta BIGINT, " +
                                    "    NoPeserta VARCHAR(50), " +
                                    "    UserId VARCHAR(100), " +
                                    "    NoPaket BIGINT, " +
                                    "    Filename VARCHAR(255), " +
                                    "    FilePath VARCHAR(500), " +
                                    "    CapturedAt DATETIME DEFAULT GETDATE(), " +
                                    "    UserInput VARCHAR(100), " +
                                    "    TimeInput DATETIME DEFAULT GETDATE() " +
                                    "  ); " +
                                    "END;"
                                cmdCreateTable.ExecuteNonQuery() |> ignore
                                
                                // Insert photo record
                                use cmdInsert = new SqlCommand()
                                cmdInsert.Connection <- conn
                                cmdInsert.CommandType <- CommandType.Text
                                cmdInsert.CommandText <-
                                    "INSERT INTO WISECON_PSIKOTEST.dbo.TR_PesertaFoto " +
                                    "(IdNoPeserta, NoPeserta, UserId, NoPaket, Filename, FilePath, CapturedAt, UserInput, TimeInput) " +
                                    "VALUES (@idNoPeserta, @noPeserta, @userId, @noPaket, @filename, @filepath, GETDATE(), @userId, GETDATE())"
                                cmdInsert.Parameters.AddWithValue("@idNoPeserta", idNoPeserta) |> ignore
                                cmdInsert.Parameters.AddWithValue("@noPeserta", noPeserta) |> ignore
                                cmdInsert.Parameters.AddWithValue("@userId", userId) |> ignore
                                cmdInsert.Parameters.AddWithValue("@noPaket", noPaket) |> ignore
                                cmdInsert.Parameters.AddWithValue("@filename", filename) |> ignore
                                cmdInsert.Parameters.AddWithValue("@filepath", fullPath) |> ignore
                                cmdInsert.ExecuteNonQuery() |> ignore
                                
                                this.Ok(box {| success = true; filename = filename |}) :> IActionResult
                        finally
                            conn.Close()
        with ex ->
            this.BadRequest(ex.Message) :> IActionResult
