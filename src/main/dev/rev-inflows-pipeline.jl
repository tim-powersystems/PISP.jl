using PISP

downloadpath       = normpath("/Volumes/Seagate/CSIRO AR-PST Stage 5/PISP-downloads")
download_from_AEMO = true
poe                = 10
reftrace           = 4006
years              = [2025]
output_root        = normpath("/Volumes/Seagate/CSIRO AR-PST Stage 5/PISP-outputs")
write_csv          = true
write_arrow        = false
scenarios          = [1,2,3]

# Pre-change run (run before applying the /n fix)
# PISP.build_ISP24_datasets(
#     downloadpath       = downloadpath,
#     download_from_AEMO = download_from_AEMO,
#     poe                = poe,
#     reftrace           = reftrace,
#     years              = years,
#     output_name        = "out-inflows-pre-n",
#     output_root        = output_root,
#     write_csv          = write_csv,
#     write_arrow        = write_arrow,
#     scenarios          = scenarios,
# )

# Post-change run (run after applying the /n fix)
PISP.build_ISP24_datasets(
    downloadpath       = downloadpath,
    download_from_AEMO = download_from_AEMO,
    poe                = poe,
    reftrace           = reftrace,
    years              = years,
    output_name        = "out-inflows-post-n",
    output_root        = output_root,
    write_csv          = write_csv,
    write_arrow        = write_arrow,
    scenarios          = scenarios,
)
