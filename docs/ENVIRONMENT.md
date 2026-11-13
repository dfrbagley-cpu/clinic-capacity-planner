# Runtime provenance and reproducibility

The project does not currently ship a dependency lockfile. CI checks current R
and available package versions; this is not a guarantee that a future installation
will reproduce an earlier environment. Do not interpret a green source check as
acceptance of a hospital workstation or clinical deployment.

Before local acceptance, run from the source checkout:

```sh
Rscript --vanilla scripts/capture_environment.R /approved/local/path/environment.csv
```

Replace the output path with a new file in an approved local folder. The helper
uses base R, installs nothing, and records R/platform and installed package
names/versions, not patient records, usernames, machine names or library paths.
It refuses to overwrite an existing file. Keep this inventory with the exact
source commit/release, test results and local acceptance record. Do not commit
local deployment inventories automatically. Review any file before sharing.

For repeatable deployment, ask IT to retain the approved R installer and package
artifacts or an approved package snapshot matching that inventory. An inventory
alone cannot restore packages and does not pin dependencies. Record desktop Excel
version separately if the workflow uses it. Re-run the synthetic tests and local
acceptance after changing R, packages or the relevant desktop software.

A future lockfile should be produced from an actually tested environment, not
invented version numbers. Existing run receipts supplement, not replace, this
local acceptance record.
