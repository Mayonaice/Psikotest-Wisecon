namespace PsikotestWisesa.Controllers

open System
open System.Data
open System.Security.Claims
open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authentication
open Microsoft.AspNetCore.Authentication.Cookies
open Microsoft.AspNetCore.Authorization

[<CLIMutable>]
type LoginRequest = { UserID: string; Password: string }

type AccountController (db: IDbConnection) =
    inherit Controller()

    [<AllowAnonymous>]
    member this.Login () : IActionResult =
        this.View()

    [<HttpPost>]
    [<AllowAnonymous>]
    [<ValidateAntiForgeryToken>]
    member this.Login (model: LoginRequest) : IActionResult =
        if isNull model.UserID || isNull model.Password || model.UserID.Trim() = "" || model.Password.Trim() = "" then
            this.ViewData.["ErrorMessage"] <- "Silakan isi User ID dan Password"
            this.View(box model)
        else
            let conn = db :?> Microsoft.Data.SqlClient.SqlConnection
            use cmd = new Microsoft.Data.SqlClient.SqlCommand()
            cmd.Connection <- conn
            cmd.CommandText <- "dbo.WEB_SP_Login"
            cmd.CommandType <- CommandType.StoredProcedure
            cmd.Parameters.AddWithValue("@UserID", model.UserID.Trim()) |> ignore
            cmd.Parameters.AddWithValue("@Password", model.Password.Trim()) |> ignore

            conn.Open()
            let result =
                try
                    try
                        use reader = cmd.ExecuteReader()
                        let mutable isOk = false
                        if reader.HasRows then
                            isOk <- true
                        reader.Close()

                        if isOk then
                            use cmd2 = new Microsoft.Data.SqlClient.SqlCommand()
                            cmd2.Connection <- conn
                            cmd2.CommandText <- "SELECT TOP 1 EmployeeCode, Name FROM DA.dbo.HRD_Employee WHERE EmployeeCode = @UserID OR Email = @UserID OR PhoneNumber = @UserID OR '0' + SUBSTRING(PhoneNumber, 3, LEN(PhoneNumber)) = @UserID"
                            cmd2.Parameters.AddWithValue("@UserID", model.UserID.Trim()) |> ignore
                            use rdr2 = cmd2.ExecuteReader()
                            let mutable empCode = model.UserID.Trim()
                            let mutable empName = model.UserID.Trim()
                            if rdr2.Read() then
                                empCode <- rdr2.GetString(0)
                                empName <- rdr2.GetString(1)
                            rdr2.Close()
                            let claims = [
                                new Claim(ClaimTypes.NameIdentifier, model.UserID.Trim())
                                new Claim(ClaimTypes.Name, model.UserID.Trim())
                                new Claim("EmployeeCode", empCode)
                                new Claim("EmployeeName", empName)
                                new Claim(ClaimTypes.GivenName, empName)
                            ]
                            let identity = new ClaimsIdentity(claims, CookieAuthenticationDefaults.AuthenticationScheme)
                            let principal = new ClaimsPrincipal(identity)
                            this.HttpContext.SignInAsync(CookieAuthenticationDefaults.AuthenticationScheme, principal) |> ignore
                            this.RedirectToAction("Applicants", "Dashboard") :> IActionResult
                        else
                            this.View(box model) :> IActionResult
                    with
                    | :? Microsoft.Data.SqlClient.SqlException as ex ->
                        this.ViewData.["ErrorMessage"] <- ex.Message
                        this.View(box model) :> IActionResult
                finally
                    conn.Close()
            result

    member this.Logout () : IActionResult =
        this.HttpContext.SignOutAsync(CookieAuthenticationDefaults.AuthenticationScheme) |> ignore
        this.RedirectToAction("Login")
