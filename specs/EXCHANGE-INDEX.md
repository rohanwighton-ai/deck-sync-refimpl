# Excel Control Layer — exchange index

The full design conversation between the Excel side (Claude AI, working through Rohan) and
the PowerPoint/VBA side (Claude Code), 31 July 2026. Filed here because **this repo is the
durable record; the OneDrive `Claude/` folder is only the transport.**

Read in order. Rounds 1–3 established the contract, 4–6 corrected it against measurement,
7 paused everything, 8–10 resolved the pause, 11–12 delivered.

| # | File | From | What it settled |
|---|---|---|---|
| 1 | `Excel_Control_Layer_Specification.docx` | Excel | The original contract: field register, one row per Quarter × EntityCode × FieldID, R1–R8 |
| 2 | `Excel_Control_Layer_Response.md` | VBA | **R1 refused on evidence** — `ShapeName` is not unique (24 duplicate names across 71 of 166 shapes on one slide). Counter-proposal: join on `FieldID` ↔ shape `role` tag |
| 3 | `Excel_Control_Layer_Confirmation.md` | Excel | R1 withdrawn. D1–D12. `Status` approval workflow accepted |
| 4 | `Excel_Control_Layer_Round4.md` | VBA | Measured answers to Q5–Q12; two amendments; the field inventory (77 items) as Appendix A |
| 5a | `Excel_Control_Layer_Round5_Addendum.md` | Excel | `SlideID` withdrawn from the register |
| 5b | `Excel_Control_Layer_Round5_Consolidated.md` | Excel | **E2 delivered** — the five-row FieldID map. Three-class taxonomy |
| 6 | `Excel_Control_Layer_Round6.md` | VBA | **Blocker found:** `EntityCode` had the same problem `FieldID` had. F1–F4 requested |
| 7 | `Excel_Control_Layer_Round7.md` | Excel | **Both lanes paused** — the compiler had never been described |
| — | `Excel_Control_Layer_Protocol1.md` | Excel | Standing protocol: lane boundaries, scope flags, RM decision categories |
| 8 | `Excel_Control_Layer_Round8.md` | VBA | The compiler rundown. The human-compiler ruling dissolved the VBA-vs-Office-JS fork |
| 9 | `Excel_Control_Layer_Round9.md` | Excel | **Creation must come out of the sync path.** RM5 closed as self-contradictory |
| 10 | `Excel_Control_Layer_Round10.md` | Excel | E1–E5 moved to the VBA lane. **F1–F4 released** (`FY26Q4`, `Approved`, `Text`, `ALL`) |
| 11 | `Excel_Control_Layer_Round11.md` | VBA | Stop on documents. `PROJECT_STATUS` is an enum, not prose. `-2` is three slides |
| 11b | `Decision_Note_Round11_1.md` | Excel | Fourth class named **Controlled**. Canonical vocabulary → prediction revised 12 → 19 |
| 12 | `Excel_Control_Layer_Round12.md` | VBA | **The result: 19 corrected, as predicted, verified on the real deck** |
| — | `ABOUT_BODY_Field_Package.md` | Excel | **R13** (human approval gate) + the second field, ready to wire in |

## Where the outcomes live

- Architecture decisions → `~/claude-brain/DECISIONS.md` (newest first)
- Current plan and status → `../WORKPLAN.md`
- The design this all extends → `deck-compiler-concept.md`

## Open at close of 31 July

**R13 must be built before `ABOUT_BODY` is wired in.** Everything else on both sides is
released; the RM decisions are all closed.
