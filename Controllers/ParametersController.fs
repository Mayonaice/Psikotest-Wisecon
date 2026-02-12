namespace PsikotestWisesa.Controllers

open System
open System.Data
open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authorization

[<CLIMutable>]
type ParameterRequest = {
    Name: string
    Remark: string
    Value: string
    Status: Nullable<int>
    Action: string
    User: string
}

[<ApiController>]
type ParametersController (db: IDbConnection) =
    inherit Controller()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Parameters/List")>]
    member this.List () : IActionResult =
        let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
        use cmd = new Microsoft.Data.SqlClient.SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <- "SELECT Name, Remark, Value, CASE WHEN ISNULL(Status_Parameter,0)=1 THEN 'Yes' ELSE 'No' END AS Name_Status, UserInput, TimeInput, UserEdit, TimeEdit FROM WISECON_PSIKOTEST.dbo.SYS_Parameter ORDER BY Name"
        conn.Open()
        try
            use rdr = cmd.ExecuteReader()
            let rows = ResizeArray<obj>()
            while rdr.Read() do
                let name = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                let remark = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                let value = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                let statusText = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                let userInput = if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                let timeInput = if rdr.IsDBNull(5) then DateTime.MinValue else rdr.GetDateTime(5)
                let userEdit = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                let timeEdit = if rdr.IsDBNull(7) then Nullable() else Nullable(rdr.GetDateTime(7))
                rows.Add(box {| name = name; remark = remark; value = value; statusText = statusText; userInput = userInput; timeInput = timeInput; userEdit = userEdit; timeEdit = timeEdit |})
            this.Ok(rows)
        finally
            conn.Close()

    [<Authorize>]
    [<HttpGet>]
    [<Route("Parameters/Get")>]
    member this.GetOne ([<FromQuery>] name: string) : IActionResult =
        if String.IsNullOrWhiteSpace(name) then this.BadRequest(box {| error = "Missing name" |}) :> IActionResult
        else
            let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
            use cmd = new Microsoft.Data.SqlClient.SqlCommand()
            cmd.Connection <- conn
            cmd.CommandType <- CommandType.Text
            cmd.CommandText <- "SELECT Name, Remark, Value, CASE WHEN ISNULL(Status_Parameter,0)=1 THEN 'Yes' ELSE 'No' END AS Name_Status, UserInput, TimeInput, UserEdit, TimeEdit FROM WISECON_PSIKOTEST.dbo.SYS_Parameter WHERE Name=@n"
            cmd.Parameters.AddWithValue("@n", name) |> ignore
            conn.Open()
            try
                use rdr = cmd.ExecuteReader()
                if rdr.Read() then
                    let name = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                    let remark = if rdr.IsDBNull(1) then "" else rdr.GetString(1)
                    let value = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                    let statusText = if rdr.IsDBNull(3) then "" else rdr.GetString(3)
                    let userInput = if rdr.IsDBNull(4) then "" else rdr.GetString(4)
                    let timeInput = if rdr.IsDBNull(5) then DateTime.MinValue else rdr.GetDateTime(5)
                    let userEdit = if rdr.IsDBNull(6) then "" else rdr.GetString(6)
                    let timeEdit = if rdr.IsDBNull(7) then Nullable() else Nullable(rdr.GetDateTime(7))
                    this.Ok(box {| name = name; remark = remark; value = value; statusText = statusText; userInput = userInput; timeInput = timeInput; userEdit = userEdit; timeEdit = timeEdit |})
                else this.NotFound()
            finally
                conn.Close()

    [<Authorize>]
    [<HttpPost>]
    [<Route("Parameters/Modify")>]
    member this.Modify ([<FromBody>] req: ParameterRequest) : IActionResult =
        let name = if isNull req.Name then "" else req.Name
        if String.IsNullOrWhiteSpace(name) then this.BadRequest(box {| error = "Missing name" |}) :> IActionResult
        else
            let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
            use cmd = new Microsoft.Data.SqlClient.SqlCommand()
            cmd.Connection <- conn
            cmd.CommandType <- CommandType.StoredProcedure
            cmd.CommandText <- if not (isNull req.Action) && req.Action.Equals("add", StringComparison.OrdinalIgnoreCase) then "WISECON_PSIKOTEST.dbo.SP_IT_SysParameter_Insert" else "WISECON_PSIKOTEST.dbo.SP_IT_SysParameter"
            cmd.Parameters.AddWithValue("@Name", name) |> ignore
            cmd.Parameters.AddWithValue("@Remark", (if isNull req.Remark then "" else req.Remark)) |> ignore
            cmd.Parameters.AddWithValue("@Value", (if isNull req.Value then "" else req.Value)) |> ignore
            let st = if req.Status.HasValue then req.Status.Value else 0
            cmd.Parameters.AddWithValue("@Status", st) |> ignore
            cmd.Parameters.AddWithValue("@User", (if isNull req.User then "" else req.User)) |> ignore
            conn.Open()
            try
                let _ = cmd.ExecuteNonQuery()
                this.Ok()
            finally
                conn.Close()