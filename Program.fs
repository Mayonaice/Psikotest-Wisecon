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

    member private this.recalculateResults(conn: SqlConnection, userId: string) =
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
            "  WHERE (R.UserId=@u OR R.NoPeserta = (SELECT TOP 1 NoPeserta FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl WHERE UserId=@u) " +
            "     OR R.IdNoPeserta = (SELECT TOP 1 ID FROM WISECON_PSIKOTEST.dbo.MS_Peserta WHERE NoPeserta = (SELECT TOP 1 NoPeserta FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl WHERE UserId=@u))) " +
            "    AND NOT EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResultHistory H WHERE H.IdNoPeserta=R.IdNoPeserta AND H.GroupSoal=R.GroupSoal AND H.AttemptTime=COALESCE(D.TimeInput, R.TimeInput)); " +
            "END;"
        insHist.Parameters.AddWithValue("@u", userId) |> ignore
        insHist.ExecuteNonQuery() |> ignore
        use del = new SqlCommand("DELETE FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult WHERE UserId=@u OR NoPeserta = (SELECT TOP 1 NoPeserta FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl WHERE UserId=@u) OR IdNoPeserta = (SELECT TOP 1 ID FROM WISECON_PSIKOTEST.dbo.MS_Peserta WHERE NoPeserta = (SELECT TOP 1 NoPeserta FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl WHERE UserId=@u))", conn)
        del.Parameters.AddWithValue("@u", userId) |> ignore
        del.ExecuteNonQuery() |> ignore
        use hitung = new SqlCommand("dbo.SP_HitungNilaiUjian", conn)
        hitung.CommandType <- CommandType.StoredProcedure
        hitung.Parameters.AddWithValue("@UserId", userId) |> ignore
        hitung.ExecuteNonQuery() |> ignore

    override this.ExecuteAsync(stoppingToken: CancellationToken) : Task =
        task {
            while not stoppingToken.IsCancellationRequested do
                try
                    let connStr = cfg.GetConnectionString("DefaultConnection")
                    use conn = new SqlConnection(connStr)
                    conn.Open()

                    use cmd = new SqlCommand()
                    cmd.Connection <- conn
                    cmd.CommandType <- CommandType.Text
                    cmd.CommandText <-
                        "SELECT TOP 50 D.UserId, D.NoPaket, D.NoPeserta " +
                        "FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl D " +
                        "WHERE D.StartTest IS NOT NULL AND DATEADD(HOUR, 1, D.StartTest) <= GETDATE() " +
                        "AND D.TimeInput = (SELECT MAX(D2.TimeInput) FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl D2 WHERE D2.NoPeserta = D.NoPeserta) " +
                        "AND EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.TR_Psikotest T WHERE T.UserId = D.UserId) " +
                        "AND NOT EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult r WHERE r.IdNoPeserta = (SELECT TOP 1 ID FROM WISECON_PSIKOTEST.dbo.MS_Peserta P WHERE P.NoPeserta = D.NoPeserta)) " +
                        "ORDER BY D.StartTest"

                    use rdr = cmd.ExecuteReader()
                    let rows = ResizeArray<string * int64 * string>()
                    while rdr.Read() do
                        let userId = if rdr.IsDBNull(0) then "" else rdr.GetString(0)
                        let noPaket = if rdr.IsDBNull(1) then 0L else Convert.ToInt64(rdr.GetValue(1))
                        let noPeserta = if rdr.IsDBNull(2) then "" else rdr.GetString(2)
                        if not (String.IsNullOrWhiteSpace(userId)) && noPaket > 0L && not (String.IsNullOrWhiteSpace(noPeserta)) then
                            rows.Add((userId, noPaket, noPeserta))
                    rdr.Close()

                    for (userId, noPaket, noPeserta) in rows do
                        try
                            use upd = new SqlCommand()
                            upd.Connection <- conn
                            upd.CommandType <- CommandType.Text
                            upd.CommandText <-
                                "UPDATE WISECON_PSIKOTEST.dbo.MS_PesertaDtl " +
                                "SET TimeEdit = GETDATE(), UserEdit='SYSTEM' " +
                                "WHERE UserId=@u AND NoPaket=@p AND NoPeserta=@n " +
                                "AND StartTest IS NOT NULL AND DATEADD(HOUR, 1, StartTest) <= GETDATE() " +
                                "AND TimeInput = (SELECT MAX(D2.TimeInput) FROM WISECON_PSIKOTEST.dbo.MS_PesertaDtl D2 WHERE D2.NoPeserta = @n) " +
                                "AND NOT EXISTS (SELECT 1 FROM WISECON_PSIKOTEST.dbo.TR_PsikotestResult r WHERE r.IdNoPeserta = (SELECT TOP 1 ID FROM WISECON_PSIKOTEST.dbo.MS_Peserta P WHERE P.NoPeserta = @n)); " +
                                "SELECT @@ROWCOUNT;"
                            upd.Parameters.AddWithValue("@u", userId) |> ignore
                            upd.Parameters.AddWithValue("@p", noPaket) |> ignore
                            upd.Parameters.AddWithValue("@n", noPeserta) |> ignore
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

                                this.recalculateResults(conn, userId)
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
