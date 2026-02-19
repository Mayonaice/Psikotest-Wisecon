namespace PsikotestWisesa.Controllers

open System
open System.Data
open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authorization

[<CLIMutable>]
type PetunjukRequest = {
    SeqNo: Nullable<int64>
    Keterangan: string
    bAktif: bool
    User: string
}

[<ApiController>]
type GuidesController (db: IDbConnection) =
    inherit Controller()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Guides/List")>]
    member this.List([<FromQuery>] start: Nullable<DateTime>, [<FromQuery>] endd: Nullable<DateTime>, [<FromQuery>] status: string) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        let mutable whereClause = " WHERE 1=1 "
        if start.HasValue && endd.HasValue then
            whereClause <- whereClause + " AND TimeInput BETWEEN @s AND @e "
            cmd.Parameters.AddWithValue("@s", start.Value.Date) |> ignore
            cmd.Parameters.AddWithValue("@e", (endd.Value.Date.AddDays(1.0).AddSeconds(-1.0))) |> ignore
        match status with
        | null -> ()
        | s when String.IsNullOrWhiteSpace(s) -> ()
        | s when s.Equals("AKTIF", StringComparison.OrdinalIgnoreCase) -> whereClause <- whereClause + " AND ISNULL(bAktif,0)=1 "
        | s when s.Equals("NONAKTIF", StringComparison.OrdinalIgnoreCase) -> whereClause <- whereClause + " AND ISNULL(bAktif,0)=0 "
        | _ -> ()
        cmd.CommandText <- "SELECT SeqNo, Keterangan, CASE WHEN bAktif=1 THEN 'AKTIF' ELSE 'NONAKTIF' END AS Status, UserInput, TimeInput, UserEdit, TimeEdit FROM WISECON_PSIKOTEST.dbo.MS_Petunjuk" + whereClause + " ORDER BY SeqNo"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let seqNo =
                    let v = rdr.GetValue(0)
                    match v with
                    | :? int64 as x -> x
                    | :? int32 as x -> int64 x
                    | :? int as x -> int64 x
                    | _ -> try Convert.ToInt64(v) with _ -> 0L
                let ket = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                let status = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                let userInput = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                let timeInput = if rdr.IsDBNull(4) then DateTime.MinValue else rdr.GetDateTime(4)
                let userEdit = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                let timeEdit = if rdr.IsDBNull(6) then DateTime.MinValue else rdr.GetDateTime(6)
                rows.Add(box {| seqNo = string seqNo; keterangan = ket; status = status; userInput = userInput; timeInput = timeInput; userEdit = userEdit; timeEdit = timeEdit |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Guides/Get")>]
    member this.GetOne([<FromQuery>] seqNo: int64) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT SeqNo, REPLACE(Keterangan, '''', '\''') Keterangan, bAktif, UserInput, TimeInput, UserEdit, TimeEdit FROM WISECON_PSIKOTEST.dbo.MS_Petunjuk WHERE SeqNo=@s"
        cmd.Parameters.AddWithValue("@s", seqNo) |> ignore
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            if rdr.Read() then
                let seqNo =
                    let v = rdr.GetValue(0)
                    match v with
                    | :? int64 as x -> x
                    | :? int32 as x -> int64 x
                    | :? int as x -> int64 x
                    | _ -> try Convert.ToInt64(v) with _ -> 0L
                let ket = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                let bAktif = if rdr.IsDBNull(2) then false else rdr.GetBoolean(2)
                let userInput = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                let timeInput = if rdr.IsDBNull(4) then DateTime.MinValue else rdr.GetDateTime(4)
                let userEdit = if rdr.IsDBNull(5) then "" else rdr.GetString(5)
                let timeEdit = if rdr.IsDBNull(6) then DateTime.MinValue else rdr.GetDateTime(6)
                this.Ok(box {| seqNo = seqNo; keterangan = ket; bAktif = bAktif; userInput = userInput; timeInput = timeInput; userEdit = userEdit; timeEdit = timeEdit |})
            else this.NotFound()
        finally
            conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Guides/Modify")>]
    member this.Modify([<FromBody>] req: PetunjukRequest) : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.StoredProcedure
        cmd.CommandText <- "dbo.SP_Petunjuk"
        cmd.CommandTimeout <- 120
        let seq = if req.SeqNo.HasValue then int req.SeqNo.Value else 0
        let ket = if isNull req.Keterangan then "" else req.Keterangan
        let pSeq = new Microsoft.Data.SqlClient.SqlParameter("@SeqNo", SqlDbType.Int)
        pSeq.Value <- seq
        let pKet = new Microsoft.Data.SqlClient.SqlParameter("@Keterangan", SqlDbType.VarChar)
        pKet.Value <- ket
        let pAktif = new Microsoft.Data.SqlClient.SqlParameter("@bAktif", SqlDbType.Bit)
        pAktif.Value <- req.bAktif
        let pUser = new Microsoft.Data.SqlClient.SqlParameter("@User", SqlDbType.VarChar)
        pUser.Value <- (if isNull req.User then "" else req.User)
        let pAct = new Microsoft.Data.SqlClient.SqlParameter("@Act", SqlDbType.VarChar)
        pAct.Value <- "ADD"
        cmd.Parameters.AddRange([| pSeq; pKet; pAktif; pUser; pAct |]) |> ignore
        conn.Open()
        try
            let _ = cmd.ExecuteNonQuery()
            this.Ok()
        finally
            conn.Close()
