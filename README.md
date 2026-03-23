Option Explicit

'=========================================================
' Builds an 8-slide consulting-style deck:
' Augusta Food Truck Tasting Route (example structure)
'=========================================================
Public Sub Build_Augusta_FoodTruck_TastingRoute_8Slides()

    Dim pres As Presentation
    Set pres = ActivePresentation

    'Start clean (optional):
    'DeleteAllSlides pres

    'Create slides
    AddSlide_Title pres, _
        "Augusta Food Truck Tasting Route", _
        "A consulting-style route plan for a curated tasting experience" & vbCrLf & _
        "Prepared for: <Client/Group Name> | Date: " & Format(Date, "mmmm d, yyyy")

    AddSlide_Agenda pres, _
        "Agenda", _
        Array( _
            "Objective & success criteria", _
            "Recommended route overview", _
            "Stops, timing, and travel logic", _
            "Menu strategy & tasting rubric", _
            "Budget and logistics", _
            "Risks & contingencies", _
            "Next steps" _
        )

    AddSlide_Objective pres, _
        "Objective & Success Criteria", _
        "Objective", "Design a time-efficient route that maximizes variety and minimizes wait/travel time.", _
        "Success criteria", _
        Array( _
            "6–8 tastings across distinct cuisines", _
            "Total route time ≤ 3 hours (excluding optional breaks)", _
            "Balanced mix of savory + sweet", _
            "Backup options for closures/long lines" _
        )

    AddSlide_RouteOverview pres, _
        "Recommended Route Overview (High Level)", _
        Array( _
            "Route style: Loop (reduces backtracking)", _
            "Total stops: 6 primary + 2 optional backups", _
            "Travel approach: Cluster by proximity; avoid peak queue windows", _
            "Tasting approach: Small portions, shareable items, ranked scoring" _
        ), _
        "Insert map screenshot here (or embed Bing Maps image)"

    AddSlide_StopsTimeline pres, _
        "Stops & Timing (Example Schedule)", _
        Array( _
            "Stop 1 (0:00–0:20) — Quick savory starter (low queue risk)", _
            "Stop 2 (0:25–0:45) — Signature entrée bite", _
            "Stop 3 (0:50–1:10) — Regional specialty", _
            "Stop 4 (1:20–1:40) — Vegetarian/seafood option", _
            "Stop 5 (1:50–2:10) — Crowd favorite / top-rated truck", _
            "Stop 6 (2:20–2:40) — Dessert / sweet finish", _
            "Optional: Stops 7–8 as backups or bonus" _
        ), _
        "Tip: Shift 15 minutes earlier/later to avoid peak lunch rush."

    AddSlide_TastingRubric pres, _
        "Menu Strategy & Tasting Rubric", _
        Array( _
            "Order strategy: 1 signature item + 1 unique item per truck", _
            "Portion plan: 1 item shared between 2–4 people", _
            "Hydration cadence: water every 2 stops; palate reset (citrus/mint)", _
            "Dietary coverage: include at least 1 vegan/veg and 1 gluten-aware option" _
        ), _
        Array( _
            "Rubric (1–5)", _
            "Taste / seasoning balance", _
            "Texture & freshness", _
            "Creativity / uniqueness", _
            "Value for portion", _
            "Speed of service" _
        )

    AddSlide_BudgetLogistics pres, _
        "Budget, Logistics & Roles", _
        Array( _
            "Budget (per person): $20–$35 typical (varies by portion sharing)", _
            "Payments: prefer card-enabled trucks; keep small cash as backup", _
            "Transportation: carpool recommended; rideshare if parking limited", _
            "Group roles: Navigator, Order lead, Score keeper, Photo/notes" _
        ), _
        Array( _
            "Operational checklist", _
            "Check truck hours/social posts day-of", _
            "Confirm parking availability at each cluster", _
            "Bring wipes/napkins, hand sanitizer", _
            "Carry water, small cooler optional" _
        )

    AddSlide_RisksNextSteps pres, _
        "Risks, Contingencies & Next Steps", _
        Array( _
            "Top risks: truck closes early, long lines, sold-out items, weather", _
            "Mitigations: 2 backup trucks; flexible timing; indoor fallback dessert stop", _
            "Decision points: if wait > 15 min, switch to next stop", _
