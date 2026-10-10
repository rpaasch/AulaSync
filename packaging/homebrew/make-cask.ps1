# Skriver Homebrew-cask'en Casks/aulasync.rb (tap'en rpaasch/homebrew-tap) for AulaSync <version>. SHA256 regnes ud fra de
# to dmg'er; URL'erne peger på udgivelsen i <owner/repo>. release.yml prøver den af på en Mac, og udgiv.ps1 skubber den til
# tap'en. Kør den med pwsh (PowerShell 7). Windows PowerShell 5.1 læser filen som ANSI, så dér kun gennem udgiv.ps1, som
# henter den som UTF-8.
#
# - uninstall: kun quit. Alt i uninstall køres også ved brew upgrade, så start ved login (LaunchAgent) og AulaSyncs
#   mapper fjernes kun med --zap (brew uninstall --zap --cask aulasync), ligesom Afinstallér AulaSync… gør. Start ved
#   login slettes som fil, ikke med launchctl: zap's launchctl beder om administratorens adgangskode, og AulaSyncs
#   LaunchAgent starter den kun ved login.
# - Ingen caveats: udgivelserne er signeret med Developer ID og notariseret (release.yml kræver det på versionsmærker), så
#   macOS spørger kun én gang, og Homebrew overfører godkendelsen ved brew upgrade (samme Team ID).
#
#   pwsh packaging/homebrew/make-cask.ps1 3.2.0 out/dmg/AulaSync-3.2.0-arm64.dmg out/dmg/AulaSync-3.2.0-x64.dmg rpaasch/AulaSync out/homebrew/aulasync.rb
param(
    [Parameter(Mandatory = $true)][string]$Version,
    [Parameter(Mandatory = $true)][string]$Arm,
    [Parameter(Mandatory = $true)][string]$Intel,
    [Parameter(Mandatory = $true)][string]$Repo,
    [Parameter(Mandatory = $true)][string]$Out
)
$ErrorActionPreference = 'Stop'

function Full($path) { $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($path) }
function Sha($path) { (Get-FileHash -Algorithm SHA256 -LiteralPath (Full $path)).Hash.ToLowerInvariant() }

if ($Version -notmatch '^\d+\.\d+\.\d+$') { throw "Ugyldig version: $Version" }
if ($Repo -notmatch '^[\w.-]+/[\w.-]+$') { throw "Ugyldigt repo: $Repo" }

$cask = @'
cask "aulasync" do
  arch arm: "arm64", intel: "x64"

  version "@VERSION@"
  sha256 arm:   "@ARM@",
         intel: "@INTEL@"

  url "https://github.com/@REPO@/releases/download/v#{version}/AulaSync-#{version}-#{arch}.dmg"
  name "AulaSync"
  desc "Skemaer fra Aula som kalendere, der holder sig opdateret"
  homepage "https://github.com/@REPO@"

  livecheck do
    url :url
    strategy :github_latest
  end

  depends_on macos: :sequoia

  app "AulaSync.app"

  uninstall quit: "dk.rpaasch.aulasync"

  zap trash: [
    "~/Library/Application Support/AulaSync",
    "~/Library/Caches/dk.rpaasch.aulasync",
    "~/Library/HTTPStorages/dk.rpaasch.aulasync",
    "~/Library/HTTPStorages/dk.rpaasch.aulasync.binarycookies",
    "~/Library/LaunchAgents/dk.rpaasch.aulasync.plist",
    "~/Library/Preferences/dk.rpaasch.aulasync.plist",
    "~/Library/Saved Application State/dk.rpaasch.aulasync.savedState",
    "~/Library/WebKit/dk.rpaasch.aulasync",
  ]
end
'@

$text = $cask.Replace('@VERSION@', $Version).Replace('@ARM@', (Sha $Arm)).Replace('@INTEL@', (Sha $Intel)).Replace('@REPO@', $Repo)
$text = ($text -replace "`r`n", "`n") + "`n"
$target = Full $Out
$dir = Split-Path -Parent $target
if ($dir) { [IO.Directory]::CreateDirectory($dir) | Out-Null }
[IO.File]::WriteAllText($target, $text, (New-Object Text.UTF8Encoding $false))
Write-Output $target
