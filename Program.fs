namespace PsikotestWisesa

#nowarn "20"

open System
open System.Collections.Generic
open System.IO
open System.Linq
open System.Threading
open System.Threading.Tasks
open Microsoft.AspNetCore
open Microsoft.AspNetCore.Builder
open Microsoft.AspNetCore.Hosting
open Microsoft.AspNetCore.HttpsPolicy
open Microsoft.Extensions.Configuration
open Microsoft.Extensions.DependencyInjection
open Microsoft.Extensions.Hosting
open Microsoft.Extensions.Logging
open Microsoft.Data.SqlClient
open System.Data
open Microsoft.AspNetCore.Authentication
open Microsoft.AspNetCore.Authentication.Cookies
open Microsoft.AspNetCore.Authorization
open Microsoft.AspNetCore.Mvc.Authorization
open Microsoft.AspNetCore.Http

type AutoFinalizeUjianService(cfg: IConfiguration) =
    inherit BackgroundService()

    member private this.getNonIdentityColumns(conn: SqlConnection, dbName: string) =
        use cmd = new SqlCommand()
        cmd.Connection <- conn
        cmd.CommandType <- CommandType.Text
        cmd.CommandText <-
            "SELECT c.name " +
            "FROM " + dbName + ".sys.columns c " +
            "WHERE c.object_id = OBJECT_ID('" + dbName + ".dbo.TR_PsikotestResult') AND c.is_identity=0"
        use rdr = cmd.ExecuteReader()
        let cols = ResizeArray<string>()
        while rdr.Read() do
            cols.Add(rdr.GetString(0))
        rdr.Close()
        cols |> Seq.toList

    member private this.syncResults(conn: SqlConnection, userId: string) =
        let wiseCols = this.getNonIdentityColumns(conn, "WISECON_PSIKOTEST") |> Set.ofList
        let advCols = this.getNonIdentityColumns(conn, "ADVPSIKOTEST") |> Set.ofList
        let cols = Set.intersect wiseCols advCols |> Set.toList
        if cols.Length > 0 then
            use del = new SqlCommand("DELETE FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult WHERE UserId=@u", conn)
            del.Parameters.AddWithValue("@u", userId) |> ignore
            del.ExecuteNonQuery() |> ignore

            let colsSql = String.Join(", ", cols |> List.map (fun c -> "[" + c + "]"))
            use ins = new SqlCommand()
            ins.Connection <- conn
            ins.CommandType <- CommandType.Text
            ins.CommandText <- "INSERT INTO WISECON_PSIKOTEST.dbo.TR_PsikotestResult (" + colsSql + ") SELECT " + colsSql + " FROM ADVPSIKOTEST.dbo.TR_PsikotestResult WHERE UserId=@u"
            ins.Parameters.AddWithValue("@u", userId) |> ignore
            ins.ExecuteNonQuery() |> ignore

    override this.ExecuteAsync(stoppingToken: CancellationToken) : Task =
        task {
            while not stoppingToken.IsCancellationRequested do
                try
                    let connStr = cfg.GetConnectionString("DefaultConnection")
                    let advConnStr = cfg.GetConnectionString("AdvPsikotestConnection")
                    use conn = new SqlConnection(connStr)
                    conn.Open()

                    use cmd = new SqlCommand()
                    cmd.Connection <- conn
                    cmd.CommandType <- CommandType.Text
                    cmd.CommandText <-
                        "SELECT TOP 50 UserId, NoPaket " +
                        "FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl D " +
                        "WHERE D.StartTest IS NOT NULL AND DATEADD(HOUR, 1, D.StartTest) <= GETDATE() " +
                        "AND NOT EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult r WHERE r.UserId = D.UserId) " +
                        "ORDER BY D.StartTest"

                    use rdr = cmd.ExecuteReader()
                    let rows = ResizeArray<string * int64>()
                    while rdr.Read() do
                        let userId = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                        let noPaket = if rdr.IsDBNull(1) then 0L else Convert.ToInt64(rdr.GetValue(1))
                        if not (String.IsNullOrWhiteSpace(userId)) && noPaket > 0L then
                            rows.Add((userId, noPaket))
                    rdr.Close()

                    for (userId, noPaket) in rows do
                        try
                            use upd = new SqlCommand()
                            upd.Connection <- conn
                            upd.CommandType <- CommandType.Text
                            upd.CommandText <-
                                "UPDATE WISECON_PSIKOTEST.dbo.MS_PesertaDtl " +
                                "SET TimeEdit = GETDATE(), UserEdit='SYSTEM' " +
                                "WHERE UserId=@u AND NoPaket=@p " +
                                "AND StartTest IS NOT NULL AND DATEADD(HOUR, 1, StartTest) <= GETDATE() " +
                                "AND NOT EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult r WHERE r.UserId = @u); " +
                                "SELECT @@ROWCOUNT;"
                            upd.Parameters.AddWithValue("@u", userId) |> ignore
                            upd.Parameters.AddWithValue("@p", noPaket) |> ignore
                            let claimed = try Convert.ToInt32(upd.ExecuteScalar()) with _ -> 0
                            if claimed > 0 then
                                use upd2 = new SqlCommand()
                                upd2.Connection <- conn
                                upd2.CommandType <- CommandType.Text
                                upd2.CommandText <-
                                    "IF COL_LENGTH('WISECON_PSIKOTEST.dbo.MS_PesertaDtl','bKirim') IS NOT NULL " +
                                    "BEGIN " +
                                    "  UPDATE WISECON_PSIKOTEST.dbo.MS_PesertaDtl SET bKirim=1, UserEdit='SYSTEM' WHERE UserId=@u AND NoPaket=@p; " +
                                    "END; " +
                                    "IF COL_LENGTH('WISECON_PSIKOTEST.dbo.MS_PesertaDtl','WaktuKirim') IS NOT NULL " +
                                    "BEGIN " +
                                    "  UPDATE WISECON_PSIKOTEST.dbo.MS_PesertaDtl SET WaktuKirim=ISNULL(WaktuKirim,GETDATE()), UserEdit='SYSTEM' WHERE UserId=@u AND NoPaket=@p; " +
                                    "END;"
                                upd2.Parameters.AddWithValue("@u", userId) |> ignore
                                upd2.Parameters.AddWithValue("@p", noPaket) |> ignore
                                upd2.ExecuteNonQuery() |> ignore

                                use delW = new SqlCommand("DELETE FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult WHERE UserId=@u", conn)
                                delW.Parameters.AddWithValue("@u", userId) |> ignore
                                delW.ExecuteNonQuery() |> ignore

                                use adv = new SqlConnection(advConnStr)
                                adv.Open()
                                use delA = new SqlCommand("DELETE FROM ADVPSIKOTEST.dbo.TR_PsikotestResult WHERE UserId=@u", adv)
                                delA.Parameters.AddWithValue("@u", userId) |> ignore
                                delA.ExecuteNonQuery() |> ignore

                                use hitung = new SqlCommand("dbo.SP_HitungNilaiUjian", adv)
                                hitung.CommandType <- CommandType.StoredProcedure
                                hitung.Parameters.AddWithValue("@UserId", userId) |> ignore
                                hitung.ExecuteNonQuery() |> ignore

                                this.syncResults(conn, userId)
                        with _ ->
                            ()
                with _ ->
                    ()

                do! Task.Delay(TimeSpan.FromMinutes(1.0), stoppingToken)
        }

module Program =
    let exitCode = 0

    [<EntryPoint>]
    let main args =
        let builder = WebApplication.CreateBuilder(args)

        builder
            .Services
            .AddControllersWithViews(fun options ->
                let policy = AuthorizationPolicyBuilder().RequireAuthenticatedUser().Build()
                options.Filters.Add(AuthorizeFilter(policy)))
            .AddRazorRuntimeCompilation()

        builder.Services.AddRazorPages()

        builder.Services
            .AddAuthentication(CookieAuthenticationDefaults.AuthenticationScheme)
            .AddCookie(fun o ->
                o.LoginPath <- "/Account/Login"
                o.LogoutPath <- "/Account/Logout"
                o.AccessDeniedPath <- "/Account/Login"
                o.SlidingExpiration <- true)

        let connStr = builder.Configuration.GetConnectionString("DefaultConnection")
        builder.Services.AddScoped<IDbConnection>(fun _ -> new SqlConnection(connStr) :> IDbConnection)
        let advConnStr = builder.Configuration.GetConnectionString("AdvPsikotestConnection")
        builder.Services.AddScoped<SqlConnection>(fun _ -> new SqlConnection(advConnStr))
        builder.Services.AddHostedService<AutoFinalizeUjianService>()

        let app = builder.Build()

        if not (builder.Environment.IsDevelopment()) then
            app.UseExceptionHandler("/Home/Error")
            app.UseHsts() |> ignore // The default HSTS value is 30 days. You may want to change this for production scenarios, see https://aka.ms/aspnetcore-hsts.

        app.UseHttpsRedirection()

        let pathBaseEnv = Environment.GetEnvironmentVariable("ASPNETCORE_PATHBASE")
        if not (String.IsNullOrWhiteSpace(pathBaseEnv)) then
            app.UsePathBase(pathBaseEnv) |> ignore

        let pathBaseCfg = if not (builder.Environment.IsDevelopment()) then builder.Configuration.["PathBase"] else null
        if not (String.IsNullOrWhiteSpace(pathBaseCfg)) then
            app.UsePathBase(pathBaseCfg) |> ignore

        ()

        app.UseStaticFiles()
        app.UseRouting()
        app.UseAuthentication()
        app.UseAuthorization()

        app.MapControllerRoute(name = "default", pattern = "{controller=Dashboard}/{action=Applicants}/{id?}")

        app.MapRazorPages()

        app.Run()

        exitCode
