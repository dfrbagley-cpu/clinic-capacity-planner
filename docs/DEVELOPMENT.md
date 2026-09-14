# Development and source control

Maintain one public `clinic-capacity-planner` repository with two officially
named versions: **Hospital R Tool** and **Interface Version**. Basic/Full are
data modes within each version, not separate forks.

| Area | Responsibility |
|---|---|
| `R/decision_support.R` | Shared, deterministic calculations and recommendation rules |
| `R/report_pipeline.R` | Original Excel import, cleaning, attribution and reporting |
| `R/workbook_support.R` | Common workbook writers and formatting |
| `run_report.R` | Configurable reporting entry point |
| `app/`, `run_app.R` | Local Shiny interface and adapters |
| `tests/` | Core rules and interface behavior |
| `scripts/package_release.R` | Reproducible hospital and source packages |

`cc_run()` accepts validated named configuration overrides, an optional password
provider function and `keep_analysis_data = TRUE`. The latter returns the minimum
cleaned fields required to rerun the decision engine. The CLI default does not
return those records. Passwords must never be command-line arguments or config
values. Shared-drive reports still get a unique output folder on every run.

The interface loads each import in a session-scoped temporary folder. Its reactive
state is inside the server function, not a global dataset. A failed import clears
the previous results so they cannot be mistaken for a new analysis. Interactive
hour changes do not repeat Excel import or the full legacy reporting pipeline.

## Checks

```r
Rscript tests/run_tests.R
Rscript tests/test_interface.R
```

Interface tests require `shiny` and the development-only `testthat` dependency.
Both use fictional records. `tests/smoke_workflows.R` runs the actual Excel import
and reporting pipeline in Basic and Full, checks interface parity and writes a
comparison workbook in a temporary directory. It is a slower integration check:

```r
Rscript tests/smoke_workflows.R
```

GitHub Actions runs these three scripts on Linux and Windows for changes to R
code, interface files, fictional inputs or the workflow. The core checks run
before packages are installed; the interface and Excel checks then use the
current R release and log their package versions. Runs need only read access,
use fictional inputs, and cancel superseded runs on the same branch.
These checks cover unprotected Excel files and Shiny server behavior. They do not
replace the browser, protected-workbook or hospital shared-drive acceptance
checks in [ACCEPTANCE.md](ACCEPTANCE.md).

Development was checked with R 4.3.3, openxlsx 4.2.5.2 and Shiny 1.8.0. The scripts do not
install packages during ordinary report or interface execution. Hospital
deployments should retain their approved, tested dependency versions.

## Build release packages

With Git and `zip` available, from a complete checkout:

```r
Rscript scripts/package_release.R
```

The hospital archive uses an explicit manifest: no app, Shiny code, Git history,
synthetic visits or development tests. The source archive includes tracked source,
documentation and fictional examples. Input, output, credentials and local runtime
folders are excluded. Both archives are written under ignored `releases/`.

## GitHub repository

The canonical repository is
[dfrbagley-cpu/clinic-capacity-planner](https://github.com/dfrbagley-cpu/clinic-capacity-planner).
It is intentionally public so reviewers and interviewers can inspect and discuss
the work. Both versions share this repository and its calculation engine.

Only code, documentation and fictional examples belong in source control. Real
clinical records, generated operational reports and credentials stay out of Git.
The public source repository does not host patient processing. The repository's
[LICENSE](../LICENSE) records the license selected when the repository was created.
Do not place access tokens in a remote URL or commit them in configuration.

## Next development priorities

1. Reconcile Basic against one completed hospital reporting period.
2. Verify the Windows launchers, password handling and shared-drive access.
3. Review the interface with a manager using fictional data, then approved local
   data; check usability, text enlargement, downloads and error recovery.
4. Validate one real Full schedule against known staffing costs and constraints.
5. Consider managed internal hosting only if it becomes a hospital requirement.
