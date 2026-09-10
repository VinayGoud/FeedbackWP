<#
.SYNOPSIS
  Creates FeedbackPrompts and FeedbackResponses lists with all required
  columns, matching exactly what FeedbackService.ts expects.

.DESCRIPTION
  Run this same script against Dev, QA, UAT, and Prod by just changing the
  -SiteUrl parameter each time - no need to manually create columns in the
  SharePoint UI per environment.

  Safe to re-run: checks if each list/field already exists before creating
  it, so running this twice against the same site won't throw errors or
  create duplicates.

.PARAMETER SiteUrl
  The full URL of the site to provision the lists on, e.g.
  https://sonatacms.sharepoint.com/sites/SmartHomeMonitoring

.EXAMPLE
  .\Provision-FeedbackLists.ps1 -SiteUrl "https://sonatacms.sharepoint.com/sites/SmartHomeMonitoring"

.NOTES
  Requires the PnP.PowerShell module:
    Install-Module -Name PnP.PowerShell -Scope CurrentUser

  Requires at least site-owner permissions on the target site to create
  lists and columns.
#>

param(
    [Parameter(Mandatory = $true)]
    [string]$SiteUrl
)

Write-Host "Connecting to $SiteUrl ..." -ForegroundColor Cyan
Connect-PnPOnline -Url $SiteUrl -Interactive

function Ensure-List {
    param(
        [string]$ListName,
        [string]$ListTemplate = "GenericList"
    )

    $existing = Get-PnPList -Identity $ListName -ErrorAction SilentlyContinue
    if ($null -eq $existing) {
        Write-Host "Creating list: $ListName" -ForegroundColor Green
        New-PnPList -Title $ListName -Template $ListTemplate | Out-Null
    }
    else {
        Write-Host "List already exists, skipping creation: $ListName" -ForegroundColor Yellow
    }
}

function Ensure-Field {
    param(
        [string]$ListName,
        [string]$InternalName,
        [string]$DisplayName,
        [string]$Type,
        [switch]$AddToDefaultView
    )

    $existingField = Get-PnPField -List $ListName -Identity $InternalName -ErrorAction SilentlyContinue
    if ($null -eq $existingField) {
        Write-Host "  Adding field '$InternalName' ($Type) to $ListName" -ForegroundColor Green
        Add-PnPField -List $ListName -InternalName $InternalName -DisplayName $DisplayName -Type $Type -AddToDefaultView:$AddToDefaultView | Out-Null
    }
    else {
        Write-Host "  Field already exists, skipping: $InternalName" -ForegroundColor Yellow
    }
}

# ---------------------------------------------------------------------------
# FeedbackPrompts
# ---------------------------------------------------------------------------
Write-Host "`n--- FeedbackPrompts ---" -ForegroundColor Cyan
Ensure-List -ListName "FeedbackPrompts"

Ensure-Field -ListName "FeedbackPrompts" -InternalName "Active" -DisplayName "Active" -Type "Boolean" -AddToDefaultView
Ensure-Field -ListName "FeedbackPrompts" -InternalName "WeekNumber" -DisplayName "WeekNumber" -Type "Text" -AddToDefaultView

# Translation columns - internal name MUST exactly match TITLE_TRANSLATION_COLUMNS
# in FeedbackService.ts, or the web part's translation lookup silently fails
# for that language (falls back to English, no error shown).
$translationColumns = @(
    @{ Internal = "Title_Spanish";              Display = "Title (Spanish)" },
    @{ Internal = "Title_German";                Display = "Title (German)" },
    @{ Internal = "Title_French";                Display = "Title (French)" },
    @{ Internal = "Title_Italian";                Display = "Title (Italian)" },
    @{ Internal = "Title_Dutch";                Display = "Title (Dutch)" },
    @{ Internal = "Title_Japanese";                Display = "Title (Japanese)" },
    @{ Internal = "Title_Korean";                Display = "Title (Korean)" },
    @{ Internal = "Title_Polish";                Display = "Title (Polish)" },
    @{ Internal = "Title_Portuguese";                Display = "Title (Portuguese)" },
    @{ Internal = "Title_Thai";                Display = "Title (Thai)" },
    @{ Internal = "Title_Chinese_Simplified";    Display = "Title (Chinese Simplified)" },
    @{ Internal = "Title_Chinese_Traditional";   Display = "Title (Chinese Traditional)" }
)

foreach ($col in $translationColumns) {
    Ensure-Field -ListName "FeedbackPrompts" -InternalName $col.Internal -DisplayName $col.Display -Type "Text"
}

# ---------------------------------------------------------------------------
# FeedbackResponses
# ---------------------------------------------------------------------------
Write-Host "`n--- FeedbackResponses ---" -ForegroundColor Cyan
Ensure-List -ListName "FeedbackResponses"

Ensure-Field -ListName "FeedbackResponses" -InternalName "Comments" -DisplayName "Comments" -Type "Note" -AddToDefaultView
Ensure-Field -ListName "FeedbackResponses" -InternalName "Like" -DisplayName "Like" -Type "Boolean" -AddToDefaultView
Ensure-Field -ListName "FeedbackResponses" -InternalName "PromptId" -DisplayName "PromptId" -Type "Number" -AddToDefaultView
Ensure-Field -ListName "FeedbackResponses" -InternalName "CommentsTranslation" -DisplayName "CommentsTranslation" -Type "Note"

Write-Host "`nDone. Ran against: $SiteUrl" -ForegroundColor Cyan
