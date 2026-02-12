namespace PsikotestWisesa.Controllers

open Microsoft.AspNetCore.Mvc
open Microsoft.AspNetCore.Authorization

type DashboardController () =
    inherit Controller()

    [<Authorize>]
    member this.Index () : IActionResult =
        this.View()

    [<Authorize>]
    member this.Applicants () : IActionResult = this.View()

    [<Authorize>]
    member this.Questions () : IActionResult = this.View()

    [<Authorize>]
    member this.Guides () : IActionResult = this.View()

    [<Authorize>]
    member this.Parameters () : IActionResult = this.View()

    [<Authorize>]
    member this.Scores () : IActionResult = this.View()

    [<Authorize>]
    member this.Screening () : IActionResult = this.View()
