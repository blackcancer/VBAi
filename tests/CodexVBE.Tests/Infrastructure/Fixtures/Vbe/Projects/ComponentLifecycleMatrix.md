# Components, lifecycle, help and procedure scalar matrix

The qualified global JSON was inventoried before the isolated batch. All tests call the production service or parser directly.

| Source | Scenarios | Required observations |
| --- | --- | --- |
| VbeProjectComponents.cs | Native default probe; owned Word and PowerPoint dispatch; status, Save and SaveAs; probe construction failure | The same supplied probe receives identity/state/save reads; one save attempt only; exact path and macro format; construction failures propagate before mutation. |
| VbeProjectComponents.Help.cs | Default adapter context zero, positive and maximum uint; native boundary failure | Native owner zero, exact CHM path, command 0/15 and UIntPtr context; handle does not certify a topic; no real help window opens. |
| VbeProjectComponents.Lifecycle.cs | Reference GUID/major/minor/broken/built-in mutations; zero-line source; protected collection; duplicate display names with distinct paths; standalone type with non-SWP path | Every reference mutation changes the collection version; protected observations omit inaccessible data; close removes exactly the path-selected project once and preserves the namesake. |
| VbaProcedureValues.cs | Blank statements, modifiers, conditional/implicit signatures, missing/duplicate/wrong signatures, malformed parentheses/return suffixes/parameters, argument arity/names, optional defaults, all scalar types and numeric boundaries, conversion fault filtering, matrices with bad rank/lower bound, non-finite/oversized scalars, returned Null and owned COM | Exact source rejection categories; required versus optional binding; exact native scalar type; no rounding or overflow success; only overflow/invalid-cast conversion failures are translated, unrelated exceptions propagate unchanged; COM is rejected without reading its members. |

The help boundary default remains HtmlHelp. The other-host dispatch factory default remains NativeOtherHostProbe; OtherHosts.cs is unchanged. Existing standalone and Excel fixtures cover their normal dispatch and failure paths. No host, macro, network connection or user window is used. The only real COM object is a private Scripting.Dictionary instance, released after its rejection is checked.

ConversionFailureValue is an IConvertible fault fixture at Convert.ToDecimal, not a replacement for Coerce or Bind. It verifies the existing exception-filter contract, including preservation of an unrelated FormatException. No production guard is removed and no production file is excluded.
## Qualified isolated result

82 Unit tests passed, zero skipped, with XPlat Code Coverage and no exclusion setting.

| Source | Lines | Branches |
| --- | --- | --- |
| VbeProjectComponents.cs | 539/539 | 854/854 |
| VbeProjectComponents.Help.cs | 19/19 | 34/34 |
| VbeProjectComponents.Lifecycle.cs | 125/125 | 198/198 |
| VbaProcedureValues.cs | 161/161 | 314/314 |

Report: artifacts/coverage/component-lifecycle-qualified/ee9bef69-d680-4f64-82e9-81a72dc8df96/coverage.cobertura.xml.

The absolute BuildOutputRoot ends in artifacts/build/component-lifecycle-qualified. Each earlier collector had a different isolated root and had ended before the next build. Filter: TestCategory=Unit and the classes VbeProjectComponentsTests, VbeProjectExcelHostTests, VbeOtherHostPersistenceTests, VbeProjectLifecycleTests, ProjectHelp, VbaProcedureValues and VbaProcedureParamArrayTests.

The mirror gate passes with 207 mirrors for 265 production files. UTF8/CRLF and git diff --check pass. This targeted measurement does not replace the global gate or actual host qualification.