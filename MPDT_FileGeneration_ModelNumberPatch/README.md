# MPDT / ACBOS File Creation

Generates MPDT and ACBOS files from the configured input workbooks, stages generated files for ProjectWise upload, creates a ProjectWise metadata workbook, and can optionally publish the files to ProjectWise.

## Local Run

```powershell
cd "<repo>\MPDT_FileGeneration_ModelNumberPatch"
python -m pip install -r requirements.txt
python main.py --step all --file-type AUTO --target-uaid2 HS2-00002CT5W HS2-00002CT38
```

By default `config/pipeline_config.json` has `publish: false`, so step6 stages files into `output_pw_<timestamp>` and writes `PW_Upload_Metadata.xlsx`, but does not upload to ProjectWise.

## SharePoint Input And Output

`config/pipeline_config.json` controls remote file sync:

- `input_location` downloads configured Input files before every run.
- `output_location` uploads generated `Output*` and `output_pw_*` files after the run.
- The SharePoint client id and secret must be supplied as environment variables, not committed in config.

Required environment variables for SharePoint app auth:

```powershell
$env:SHAREPOINT_CLIENT_ID = "<client-id>"
$env:SHAREPOINT_CLIENT_SECRET = "<client-secret>"
```

The config fields `client_id_env` and `client_secret_env` should contain the environment variable names above, not the secret values.

## ProjectWise Preflight

Run this on the Jenkins Windows agent before enabling live upload:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\Scripts\Test-PWPSRuntime.ps1 `
  -ProjectWiseBin "C:\Program Files\Bentley\ProjectWise\bin" `
  -FailOnError
```

The preflight must pass these checks:

- 64-bit PowerShell process.
- ProjectWise Explorer/client runtime installed.
- `dmscli.dll` exists and can be loaded with its native dependencies.
- `pwps_dab` PowerShell module imports successfully.

If `dmscli.dll` exists but fails with Win32 error 126, the machine is missing a native ProjectWise dependency or the ProjectWise bin folder is not available to the running process.

## Jenkins Setup

The included `Jenkinsfile` expects a Windows agent labeled `windows-projectwise`.

Install on the Jenkins agent:

- 64-bit Python available on `PATH`.
- ProjectWise Explorer/client runtime.
- PowerShell module `pwps_dab` for the user running the Jenkins service/agent.
- Network access to SharePoint and ProjectWise.

Create these Jenkins string credentials:

- `sharepoint-client-id`
- `sharepoint-client-secret`
- `pw-password`

Jenkins parameters:

- `TARGET_UAID2`: one or more UAID_2 values, separated by commas or spaces.
- `FILE_TYPE`: use `AUTO` for the normal MPDT/ACBOS decision logic.
- `STEP`: use `all` for the full pipeline.
- `PUBLISH_PROJECTWISE`: unchecked means stage only; checked enables live ProjectWise upload.
- `PROJECTWISE_BIN`: ProjectWise bin folder on the Jenkins agent.

For the first Jenkins test, run with `PUBLISH_PROJECTWISE` unchecked and confirm the archived `output_pw_*` folder contains the generated file and `PW_Upload_Metadata.xlsx`. Then enable `PUBLISH_PROJECTWISE` for the live upload test.

## Jenkins Runner Command

The Jenkinsfile calls:

```powershell
.\Scripts\Run-JenkinsPipeline.ps1 `
  -TargetUAID2 "HS2-00002CT5W HS2-00002CT38" `
  -FileType AUTO `
  -Step all `
  -ProjectWiseBin "C:\Program Files\Bentley\ProjectWise\bin" `
  -InstallDependencies
```

For live ProjectWise upload, add `-PublishProjectWise`. The runner sets `MPDT_PUBLISH=true` only for live upload, so local config can safely keep `publish: false`.
