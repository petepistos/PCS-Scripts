<#
.SYNOPSIS
    Pistos Compliance Sentinel - New Client Onboarding Script
    Creates all required records in SmartSuite for a new PCS client.

.DESCRIPTION
    Execution order:
      1.  Look up Customer record by name, capture RecID
      2.  Create Customer Annual Metrics (CAM) record
      3.  Create Customer Receipts and Billing (CRB) record
      4.  Create Recovery and Inventory Plans record
      5.  Create 50 Risk Assessment records linked to CAM
      6.  PATCH CAM with CRB link and all 50 Risk Assessment links
      7.  PATCH all 15 Security Policy records to add customer link
      8.  PATCH all 61 Training and Awareness Video records to add customer link
      9.  PATCH Customer record with all links

.PARAMETER CustomerName
    Must exactly match the Customer record Title field in SmartSuite.

.PARAMETER AssessmentYear
    Assessment year. Defaults to 2025.

.PARAMETER ApiKey
    SmartSuite API key (40-character hex string).

.EXAMPLE
    .\New-PCSClient.ps1 -CustomerName "Apex Brokerage" -ApiKey "your_api_key_here"

.NOTES
    Run from C:\pistos\
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$CustomerName,

    [string]$AssessmentYear = "2025",

    [Parameter(Mandatory)]
    [string]$ApiKey
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# =============================================================================
# APP IDs
# =============================================================================

$AccountId = "sgca6mtm"
$BaseUrl   = "https://app.smartsuite.com/api/v1/applications"

$AppId = @{
    Customers        = "6900c72a20ac7c93267a05ed"
    CAM              = "696daa91a6798b8cec2a2757"
    CRB              = "6931b23768d032eeec052ee9"
    RiskAssessments  = "69658acc65dac0e165ad24c8"
    RecoveryPlans    = "696576e611009c1f6d6f4b81"
    SecurityPolicies = "6931b4bb6a660139aa0d6692"
    TrainingVideos   = "6931c10f3bce816c6f5c230e"
    NYCRR            = "699c7d0933c4fd456b2cdac6"
}

# =============================================================================
# FIELD SLUGS (confirmed from API)
# =============================================================================

$F = @{
    # Customers (all confirmed from API)
    Cust_Title          = "title"
    Cust_CAM            = "shatf2oq"
    Cust_CRB            = "sk5iewd5"
    Cust_RecoveryPlans  = "se5f1a15f6"
    Cust_RiskAssess     = "s40zjaz2"
    Cust_SecPolicies    = "sjvb15al"
    Cust_TrainVideos    = "s6a41a897e"

    # CAM (all confirmed from API)
    CAM_Title           = "title"
    CAM_Customer        = "s66338ee26"
    CAM_CRB             = "sjq216p5"
    CAM_RiskAssess      = "s6n9jtk9"
    CAM_AssessYear      = "sa7da430dc"

    # CRB (all confirmed from API)
    CRB_Title           = "title"
    CRB_Customer        = "sa87f7e561"
    CRB_AssessYear      = "s129d8489c"

    # Recovery Plans (all confirmed from API)
    RP_Title            = "title"
    RP_Customer         = "s7bfc29230"
    RP_OCP_Template     = "sac8394f24"
    RP_User_Guide       = "sn3ojeg1"

    # Risk Assessments (confirmed from Get-RAFieldSlugs.ps1)
    RA_Title            = "title"
    RA_ControlNumber    = "s6e4b09215"
    RA_ControlName      = "s2ca6be6cb"
    RA_InherentRisk     = "s45d89188b"
    RA_Platform         = "s6bd84614f"
    RA_Applicable       = "scb215244e"
    RA_EvidenceStatus   = "sb4a4b9b71"
    RA_ImplGuide        = "s935240aaa"
    RA_Domain           = "s20jaedo"
    RA_AssessYear       = "s1c03507e0"
    RA_LinkCAM          = "s5a5c0cdc2"
    RA_LinkPolicies     = "s57a6a29cb"

    # Security Policies
    SP_Customers        = "s6335f6ddd"

    # Training Videos
    TV_Customers        = "sa4df46192"
}

# =============================================================================
# STATIC RecID ARRAYS
# =============================================================================

$SecurityPolicyRecIDs = @(
    "699d1e62306fc9aaa33057ba",
    "699c68284b668835aae4abcf",
    "699c68284b668835aae4abd4",
    "699c70649c55ae0879ead3fc",
    "699c68284b668835aae4abcd",
    "699c68284b668835aae4abd6",
    "699c719f3bbbbb9c29edbbda",
    "699c68284b668835aae4abd0",
    "699c68284b668835aae4abd1",
    "699c6c7f9c55ae0879ead3f9",
    "699c68284b668835aae4abd3",
    "699c68284b668835aae4abd2",
    "699c68284b668835aae4abce",
    "699c68284b668835aae4abd5",
    "69c917978fe4bf681ad41990"
)

$TrainingVideoRecIDs = @(
    "6931c3b826889575bdd457ee",
    "6931c3b826889575bdd457ef",
    "6931c3b826889575bdd457f0",
    "6931c3b826889575bdd457f1",
    "6931c3b826889575bdd457f2",
    "6931c3b826889575bdd457f3",
    "6931c3b826889575bdd457f4",
    "6931c3b826889575bdd457f5",
    "6931c3b826889575bdd457f6",
    "6931c3b826889575bdd457f7",
    "6931c3b826889575bdd457f8",
    "6931c3b826889575bdd457f9",
    "6931c3b826889575bdd457fa",
    "6931c3b826889575bdd457fb",
    "6931c3b826889575bdd457fc",
    "6931c3b826889575bdd457fd",
    "6931c3b826889575bdd457fe",
    "6931c3b826889575bdd457ff",
    "6931c3b826889575bdd45800",
    "6931c3b826889575bdd45801",
    "6931c3b826889575bdd45802",
    "6931c3b826889575bdd45803",
    "6931c3b826889575bdd45804",
    "6931c3b826889575bdd45805",
    "6931c7b40baefadbec68656a",
    "6931c7b40baefadbec68656b",
    "6931c7b40baefadbec68656c",
    "6931c7b40baefadbec68656d",
    "6931c7b40baefadbec68656e",
    "6931c7b40baefadbec68656f",
    "6931c7b40baefadbec686570",
    "6931c7b40baefadbec686571",
    "6931c7b40baefadbec686572",
    "6931c7b40baefadbec686573",
    "6931c7b40baefadbec686574",
    "6931c7b40baefadbec686575",
    "6931c7b40baefadbec686576",
    "6931c7b40baefadbec686577",
    "6931c7b40baefadbec686578",
    "6931c7b40baefadbec686579",
    "6931c7b40baefadbec68657a",
    "6931c7b40baefadbec68657b",
    "6931c7b40baefadbec68657c",
    "6931c7b40baefadbec68657d",
    "6931c7b40baefadbec68657e",
    "6931c7b40baefadbec68657f",
    "6931c7b40baefadbec686580",
    "6931c7b40baefadbec686581",
    "6931c7b40baefadbec686582",
    "6931c7b40baefadbec686583",
    "6931c7b40baefadbec686584",
    "6931c7b40baefadbec686585",
    "6931c7b40baefadbec686586",
    "6931c7b40baefadbec686587",
    "6931c7b40baefadbec686588",
    "6931c7b40baefadbec686589",
    "6931c7b40baefadbec68658a",
    "698df592f5cfce2b77ba8813",
    "698df64d6b6914088ab96b33",
    "698e099b57c688b009efdec8",
    "698e09fccc61f9c157c74c60"
)

# Policy Num → RecID map (index matches Policy Num 0–14)
# Order matches $SecurityPolicyRecIDs array order
$PolicyRecIDMap = @{
    0  = "699d1e62306fc9aaa33057ba"   # Asset Management Policy
    1  = "699c68284b668835aae4abcf"   # Access Control & Identity Management Policy
    2  = "699c68284b668835aae4abd4"   # Business Continuity & Disaster Recovery Policy
    3  = "699c70649c55ae0879ead3fc"   # Customer & Consumer Privacy Policy
    4  = "699c68284b668835aae4abcd"   # Cybersecurity Governance & Program Policy
    5  = "699c68284b668835aae4abd6"   # Data Governance, Retention & Destruction Policy
    6  = "699c719f3bbbbb9c29edbbda"   # Email Authentication and Anti-Spoofing Policy
    7  = "699c68284b668835aae4abd0"   # Encryption & Data Protection Policy
    8  = "699c68284b668835aae4abd1"   # Endpoint Security & Firewall Policy
    9  = "699c6c7f9c55ae0879ead3f9"   # Human Resources Policy
    10 = "699c68284b668835aae4abd3"   # Incident Response & Breach Notification Policy
    11 = "699c68284b668835aae4abd2"   # Logging & Monitoring Policy
    12 = "699c68284b668835aae4abce"   # Risk Assessment Policy
    13 = "699c68284b668835aae4abd5"   # Vendor & Third-Party Risk Management Policy
    14 = "69c917978fe4bf681ad41990"   # Environmental and Physical Security Policy
}

function Get-PolicyRecIDs {
    param([int[]]$PolicyNums)
    return $PolicyNums | ForEach-Object { $PolicyRecIDMap[$_] }
}
# All 50 controls with full static content.
# ImplGuide is a URL field — value must be @{url=...; label=...} or $null.
# EvidenceStatus and Applicable must exactly match SmartSuite single-select options.
# =============================================================================

$EFT = "https://github.com/petepistos/skopein-scripts/tree/main/Evidence%20Form%20Templates"
$REM_W   = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/Windows%20Workstation"
$REM_M   = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/MS%20365"
$REM_AMS = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/AMS"
$REM_EXT = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/External%20Scans"

function MakeUrl($url, $label) {
    if ([string]::IsNullOrEmpty($url)) { return $null }
    return $url
}

$Controls = @(
    # ── P Controls ───────────────────────────────────────────────────────────
    @{
        ControlNumber  = "P-01"
        ControlName    = "A documented Disaster Recovery Plan is maintained and tested on a defined schedule."
        Domain         = "Operational Continuity"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 3
        Policies       = @(2)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-01_Disaster_Recovery_Plan_Reference.pdf" "P-01 Implementation Guide"
    },
    @{
        ControlNumber  = "P-02"
        ControlName    = "A documented Incident Response Plan is maintained, including roles, escalation paths, and response procedures."
        Domain         = "Operational Continuity"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 5
        Policies       = @(10)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-02_Incident_Response_Plan_Reference.pdf" "P-02 Implementation Guide"
    },
    @{
        ControlNumber  = "P-03"
        ControlName    = "A formally documented patch management process is maintained, with defined SLAs based on severity."
        Domain         = "Asset Management"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 3
        Policies       = @(0, 8)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-03_Patch_Management_Reference.pdf" "P-03 Implementation Guide"
    },
    @{
        ControlNumber  = "P-04"
        ControlName    = "A published privacy notice is maintained and includes an opt-out mechanism where applicable."
        Domain         = "Customer and Consumer Privacy"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 1
        Policies       = @(3)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-04_Privacy_Notice_OptOut.pdf" "P-04 Implementation Guide"
    },
    @{
        ControlNumber  = "P-05"
        ControlName    = "A vendor due diligence process is implemented, including annual reviews of third-party service providers."
        Domain         = "Risk Management"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 1
        Policies       = @(13)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-05_Vendor_Due_Diligence_Reference.pdf" "P-05 Implementation Guide"
    },
    @{
        ControlNumber  = "P-06"
        ControlName    = "Employees are responsible for securing laptops and mobile devices when unattended and lock screens when stepping away."
        Domain         = "Human Resources"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 2
        Policies       = @(8, 14)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-06_Device_ScreenLock_Acknowledgment.pdf" "P-06 Implementation Guide"
    },
    @{
        ControlNumber  = "P-07"
        ControlName    = "Employee onboarding and offboarding processes are documented."
        Domain         = "Human Resources"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 1
        Policies       = @(9)
        ImplGuide      = MakeUrl $EFT "P-07 Implementation Guide"
    },
    @{
        ControlNumber  = "P-08"
        ControlName    = "Only authorized personnel can access company facilities based on defined job roles and business need."
        Domain         = "Environmental Security"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 1
        Policies       = @(14)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-08_Physical_Access_Controls_Reference.pdf" "P-08 Implementation Guide"
    },
    @{
        ControlNumber  = "P-09"
        ControlName    = "Pre-employment background checks are completed for all hires, consistent with legal requirements."
        Domain         = "Human Resources"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 2
        Policies       = @(9)
        ImplGuide      = MakeUrl $EFT "P-09 Implementation Guide"
    },
    @{
        ControlNumber  = "P-10"
        ControlName    = "Remote access to internal systems requires an approved VPN with MFA."
        Domain         = "Identity and Access Management"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 5
        Policies       = @(1)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-10_Remote_Access_Request_Form.pdf" "P-10 Implementation Guide"
    },
    @{
        ControlNumber  = "P-12"
        ControlName    = "Secure configuration standards are documented and applied consistently to systems and endpoints."
        Domain         = "Identity and Access Management"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 1
        Policies       = @(0, 8)
        ImplGuide      = MakeUrl $EFT "P-12 Implementation Guide"
    },
    @{
        ControlNumber  = "P-13"
        ControlName    = "Security awareness training is required and tailored to each user's role and technical responsibilities."
        Domain         = "Human Resources"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 4
        Policies       = @(9)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-13_Security_Awareness_Training_Reference.pdf" "P-13 Implementation Guide"
    },
    @{
        ControlNumber  = "P-15"
        ControlName    = "Security policies addressing each applicable regulatory requirement have been documented, reviewed, and formally approved by management."
        Domain         = "Governance & Program"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 5
        Policies       = @(4)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-15_Security_Policy_Documentation_Reference.pdf" "P-15 Implementation Guide"
    },
    @{
        ControlNumber  = "P-16"
        ControlName    = "The entity is paperless."
        Domain         = "Customer and Consumer Privacy"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 3
        Policies       = @(5)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-16_Paperless_Office_Reference.pdf" "P-16 Implementation Guide"
    },
    @{
        ControlNumber  = "P-17"
        ControlName    = "The entity maintains cyberinsurance coverage appropriate to its risk profile and contractual obligations."
        Domain         = "Operational Continuity"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(4)
        ImplGuide      = $null
    },
    @{
        ControlNumber  = "P-18"
        ControlName    = "The entity's risk management program performs annual risk assessments that assess compliance with the NY DFS Cybersecurity regulation."
        Domain         = "Risk Management"
        Applicable     = "Applicable"
        EvidenceStatus = "Accepted"
        InherentRisk   = 5
        Policies       = @(12)
        ImplGuide      = MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-18_Annual_Risk_Assessment_Reference.pdf" "P-18 Implementation Guide"
    },
    @{
        ControlNumber  = "P-19"
        ControlName    = "Regular phishing simulations are conducted."
        Domain         = "Human Resources"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 6
        Policies       = @(9)
        ImplGuide      = $null
    },

    # ── W Controls ───────────────────────────────────────────────────────────
    @{
        ControlNumber  = "W-01"
        ControlName    = "Endpoint protection is deployed and actively maintained on all workstations."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 3
        Policies       = @(8)
        ImplGuide      = MakeUrl $REM_W "W-01 Implementation Guide"
    },
    @{
        ControlNumber  = "W-02"
        ControlName    = "Full disk encryption is enabled on all workstations."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(7, 8)
        ImplGuide      = MakeUrl $REM_W "W-02 Implementation Guide"
    },
    @{
        ControlNumber  = "W-03"
        ControlName    = "Operating system patches are applied within defined SLA timeframes based on severity."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 3
        Policies       = @(0, 8)
        ImplGuide      = MakeUrl $REM_W "W-03 Implementation Guide"
    },
    @{
        ControlNumber  = "W-04"
        ControlName    = "Local administrator accounts are disabled or restricted on all workstations."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(1, 8)
        ImplGuide      = MakeUrl $REM_W "W-04 Implementation Guide"
    },
    @{
        ControlNumber  = "W-05"
        ControlName    = "Screen lock and idle timeout policies are enforced on all workstations."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(8)
        ImplGuide      = MakeUrl $REM_W "W-05 Implementation Guide"
    },
    @{
        ControlNumber  = "W-06"
        ControlName    = "Host-based firewall is enabled and configured on all workstations."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(8)
        ImplGuide      = MakeUrl $REM_W "W-06 Implementation Guide"
    },
    @{
        ControlNumber  = "W-07"
        ControlName    = "Removable media usage is restricted or controlled by policy and technical enforcement."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(8)
        ImplGuide      = MakeUrl $REM_W "W-07 Implementation Guide"
    },
    @{
        ControlNumber  = "W-08"
        ControlName    = "Only approved software is permitted to execute on workstations."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(0, 8)
        ImplGuide      = MakeUrl $REM_W "W-08 Implementation Guide"
    },
    @{
        ControlNumber  = "W-09"
        ControlName    = "Audit logging is enabled and retained on all workstations."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(11)
        ImplGuide      = MakeUrl $REM_W "W-09 Implementation Guide"
    },
    @{
        ControlNumber  = "W-10"
        ControlName    = "Secure boot and BIOS/UEFI password protections are configured on all workstations."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(8)
        ImplGuide      = MakeUrl $REM_W "W-10 Implementation Guide"
    },
    @{
        ControlNumber  = "W-11"
        ControlName    = "Workstations are enrolled in a centralized device management platform."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(0)
        ImplGuide      = MakeUrl $REM_W "W-11 Implementation Guide"
    },

    # ── M Controls ───────────────────────────────────────────────────────────
    @{
        ControlNumber  = "M-01"
        ControlName    = "MFA is enforced for all Microsoft 365 user accounts."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 3
        Policies       = @(1)
        ImplGuide      = MakeUrl $REM_M "M-01 Implementation Guide"
    },
    @{
        ControlNumber  = "M-02"
        ControlName    = "Conditional Access policies are configured to restrict access based on risk and location."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 3
        Policies       = @(1)
        ImplGuide      = MakeUrl $REM_M "M-02 Implementation Guide"
    },
    @{
        ControlNumber  = "M-03"
        ControlName    = "Privileged administrative roles are separated from standard user accounts."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(1)
        ImplGuide      = MakeUrl $REM_M "M-03 Implementation Guide"
    },
    @{
        ControlNumber  = "M-04"
        ControlName    = "Mailbox audit logging is enabled for all Microsoft 365 users."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(11)
        ImplGuide      = MakeUrl $REM_M "M-04 Implementation Guide"
    },
    @{
        ControlNumber  = "M-05"
        ControlName    = "Anti-phishing and anti-spam policies are configured and enforced in Microsoft 365."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(6, 8)
        ImplGuide      = MakeUrl $REM_M "M-05 Implementation Guide"
    },
    @{
        ControlNumber  = "M-06"
        ControlName    = "Safe Links and Safe Attachments (Defender for Office 365) are enabled for all users."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(8)
        ImplGuide      = MakeUrl $REM_M "M-06 Implementation Guide"
    },
    @{
        ControlNumber  = "M-07"
        ControlName    = "Data Loss Prevention (DLP) policies are defined and actively enforced."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(5)
        ImplGuide      = MakeUrl $REM_M "M-07 Implementation Guide"
    },
    @{
        ControlNumber  = "M-08"
        ControlName    = "External sharing for SharePoint and OneDrive is restricted to authorized use cases."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(1, 5)
        ImplGuide      = MakeUrl $REM_M "M-08 Implementation Guide"
    },
    @{
        ControlNumber  = "M-09"
        ControlName    = "The Unified Audit Log is enabled and retained per policy."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(11)
        ImplGuide      = MakeUrl $REM_M "M-09 Implementation Guide"
    },
    @{
        ControlNumber  = "M-10"
        ControlName    = "Legacy authentication protocols are blocked for all Microsoft 365 accounts."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(1)
        ImplGuide      = MakeUrl $REM_M "M-10 Implementation Guide"
    },
    @{
        ControlNumber  = "M-11"
        ControlName    = "Identity Protection risky sign-in alerts are configured and actively monitored."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(1, 11)
        ImplGuide      = MakeUrl $REM_M "M-11 Implementation Guide"
    },

    # ── AMS Controls ─────────────────────────────────────────────────────────
    @{
        ControlNumber  = "AMS-01"
        ControlName    = "Smart Lockout is configured to prevent brute-force attacks on user accounts."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(1)
        ImplGuide      = MakeUrl $REM_AMS "AMS-01 Implementation Guide"
    },
    @{
        ControlNumber  = "AMS-02"
        ControlName    = "User account lifecycle management processes are documented and enforced."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(1)
        ImplGuide      = MakeUrl $REM_AMS "AMS-02 Implementation Guide"
    },
    @{
        ControlNumber  = "AMS-03"
        ControlName    = "Audit logging is enabled and retained for all identity and access events."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(1, 11)
        ImplGuide      = MakeUrl $REM_AMS "AMS-03 Implementation Guide"
    },
    @{
        ControlNumber  = "AMS-04"
        ControlName    = "MFA is enforced for all user accounts within the AMS environment."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(1)
        ImplGuide      = MakeUrl $REM_AMS "AMS-04 Implementation Guide"
    },
    @{
        ControlNumber  = "AMS-05"
        ControlName    = "Administrative accounts are restricted to privileged tasks only and reviewed periodically."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 2
        Policies       = @(1)
        ImplGuide      = MakeUrl $REM_AMS "AMS-05 Implementation Guide"
    },
    @{
        ControlNumber  = "AMS-06"
        ControlName    = "Data encryption is enforced at rest and in transit."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(7)
        ImplGuide      = MakeUrl $REM_AMS "AMS-06 Implementation Guide"
    },

    # ── EXT Controls ─────────────────────────────────────────────────────────
    @{
        ControlNumber  = "EXT-01"
        ControlName    = "A CAA DNS record is published to restrict unauthorized certificate issuance."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(6)
        ImplGuide      = MakeUrl $REM_EXT "EXT-01 Implementation Guide"
    },
    @{
        ControlNumber  = "EXT-02"
        ControlName    = "A DMARC policy is published and enforced to prevent email domain spoofing."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(6)
        ImplGuide      = MakeUrl $REM_EXT "EXT-02 Implementation Guide"
    },
    @{
        ControlNumber  = "EXT-03"
        ControlName    = "The robots.txt file is configured to prevent unintended exposure of sensitive paths."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(4)
        ImplGuide      = MakeUrl $REM_EXT "EXT-03 Implementation Guide"
    },
    @{
        ControlNumber  = "EXT-04"
        ControlName    = "Security headers (CSP, HSTS, X-Frame-Options, etc.) are configured on all web properties."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(4, 8)
        ImplGuide      = MakeUrl $REM_EXT "EXT-04 Implementation Guide"
    },
    @{
        ControlNumber  = "EXT-05"
        ControlName    = "An SPF record is published to authorize legitimate mail senders for the domain."
        Domain         = "Technology"
        Applicable     = "Applicable"
        EvidenceStatus = "Not Started"
        InherentRisk   = 1
        Policies       = @(6)
        ImplGuide      = MakeUrl $REM_EXT "EXT-05 Implementation Guide"
    }
)

# =============================================================================
# HELPERS
# =============================================================================

$Headers = @{
    "Authorization" = "Token $ApiKey"
    "Account-Id"    = $AccountId
    "Content-Type"  = "application/json"
}

function Invoke-SS {
    param(
        [string]$Method,
        [string]$Url,
        [hashtable]$Body = $null
    )
    $params = @{
        Method  = $Method
        Uri     = $Url
        Headers = $Headers
    }
    if ($Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress)
    }
    $result = Invoke-RestMethod @params
    Start-Sleep -Milliseconds 500
    return $result
}

function New-Record {
    param([string]$AppId, [hashtable]$Fields)
    $url = "$BaseUrl/$AppId/records/"
    return Invoke-SS -Method POST -Url $url -Body $Fields
}

function Update-Record {
    param([string]$AppId, [string]$RecId, [hashtable]$Fields)
    $url = "$BaseUrl/$AppId/records/$RecId/"
    return Invoke-SS -Method PATCH -Url $url -Body $Fields
}

function Find-CustomerRecId {
    param([string]$Name)
    $url  = "$BaseUrl/$($AppId.Customers)/records/list/"
    $body = @{ filter = @{}; sort = @(); fields_to_search = @() }
    $resp = Invoke-SS -Method POST -Url $url -Body $body
    $match = $resp.items | Where-Object { $_.title -eq $Name }
    if (-not $match) { throw "Customer '$Name' not found in SmartSuite." }
    return $match.id
}

# =============================================================================
# MAIN
# =============================================================================

Write-Host ""
Write-Host ">> Looking up Customer: $CustomerName"
$CustomerRecId = Find-CustomerRecId -Name $CustomerName
Write-Host "   OK: Customer RecID: $CustomerRecId"

# ── Step 2: CAM ──────────────────────────────────────────────────────────────
Write-Host ""
Write-Host ">> Creating Customer Annual Metrics record"
$camBody = @{
    $F.CAM_Title      = $CustomerName
    $F.CAM_Customer   = @($CustomerRecId)
    $F.CAM_AssessYear = $AssessmentYear
}
$camResp  = New-Record -AppId $AppId.CAM -Fields $camBody
$CAMRecId = $camResp.id
Write-Host "   OK: CAM RecID: $CAMRecId"

# ── Step 3: CRB ──────────────────────────────────────────────────────────────
Write-Host ""
Write-Host ">> Creating Customer Receipts and Billing record"
$crbBody = @{
    $F.CRB_Title      = $CustomerName
    $F.CRB_Customer   = @($CustomerRecId)
    $F.CRB_AssessYear = $AssessmentYear
}
$crbResp  = New-Record -AppId $AppId.CRB -Fields $crbBody
$CRBRecId = $crbResp.id
Write-Host "   OK: CRB RecID: $CRBRecId"

# ── Step 4: Recovery Plans ───────────────────────────────────────────────────
Write-Host ""
Write-Host ">> Creating Recovery and Inventory Plans record"
$rpBody = @{
    $F.RP_Title       = $CustomerName
    $F.RP_Customer    = @($CustomerRecId)
    $F.RP_OCP_Template = "https://github.com/petepistos/skopein-scripts/blob/main/Recovery-Plans/IRP_DRP_Plan.xlsx"
    $F.RP_User_Guide   = "https://github.com/petepistos/skopein-scripts/blob/main/Recovery-Plans/IRP_DRP_Plan_User_Guide.pdf"
}
$rpResp       = New-Record -AppId $AppId.RecoveryPlans -Fields $rpBody
$RPRecId      = $rpResp.id
Write-Host "   OK: Recovery Plans RecID: $RPRecId"

# ── Step 5: Risk Assessments ─────────────────────────────────────────────────
Write-Host ""
Write-Host ">> Creating $($Controls.Count) Risk Assessment records"
$RARecIds = [System.Collections.Generic.List[string]]::new()

foreach ($ctrl in $Controls) {
    $raBody = @{
        $F.RA_Title          = "$CustomerName - $($ctrl.ControlNumber)"
        $F.RA_LinkCAM        = @($CAMRecId)
        $F.RA_ControlNumber  = $ctrl.ControlNumber
        $F.RA_ControlName    = $ctrl.ControlName
        $F.RA_Domain         = $ctrl.Domain
        $F.RA_Applicable     = $ctrl.Applicable
        $F.RA_EvidenceStatus = $ctrl.EvidenceStatus
        $F.RA_InherentRisk   = $ctrl.InherentRisk
        $F.RA_AssessYear     = $AssessmentYear
    }
    if ($null -ne $ctrl.ImplGuide) {
        $raBody[$F.RA_ImplGuide] = $ctrl.ImplGuide
    }
    if ($ctrl.Policies -and $ctrl.Policies.Count -gt 0) {
        $raBody[$F.RA_LinkPolicies] = @(Get-PolicyRecIDs -PolicyNums $ctrl.Policies)
    }

    $raResp = New-Record -AppId $AppId.RiskAssessments -Fields $raBody
    $RARecIds.Add($raResp.id)
    Write-Host "   OK: $($ctrl.ControlNumber) created"
}

$total = $Controls.Count
Write-Host "   OK: $total of $total Risk Assessment records created"

# ── Step 6: Update CAM ───────────────────────────────────────────────────────
Write-Host ""
Write-Host ">> Updating CAM with CRB and Risk Assessment links"
$camPatch = @{
    $F.CAM_CRB        = @($CRBRecId)
    $F.CAM_RiskAssess = $RARecIds.ToArray()
}
Update-Record -AppId $AppId.CAM -RecId $CAMRecId -Fields $camPatch | Out-Null
Write-Host "   OK: CAM updated"

# ── Step 7: Security Policies ────────────────────────────────────────────────
Write-Host ""
Write-Host ">> Adding customer to all $($SecurityPolicyRecIDs.Count) Security Policy records"
foreach ($spId in $SecurityPolicyRecIDs) {
    $spPatch = @{ $F.SP_Customers = @($CustomerRecId) }
    Update-Record -AppId $AppId.SecurityPolicies -RecId $spId -Fields $spPatch | Out-Null
    Write-Host "   OK: Policy $spId updated"
}

# ── Step 8: Training Videos ──────────────────────────────────────────────────
Write-Host ""
Write-Host ">> Adding customer to all $($TrainingVideoRecIDs.Count) Training and Awareness Video records"
foreach ($tvId in $TrainingVideoRecIDs) {
    $tvPatch = @{ $F.TV_Customers = @($CustomerRecId) }
    Update-Record -AppId $AppId.TrainingVideos -RecId $tvId -Fields $tvPatch | Out-Null
    Write-Host "   OK: Video $tvId updated"
}

# ── Step 9: Update Customer ──────────────────────────────────────────────────
Write-Host ""
Write-Host ">> Updating Customer record with all links"
$custPatch = @{
    $F.Cust_CAM           = @($CAMRecId)
    $F.Cust_CRB           = @($CRBRecId)
    $F.Cust_RecoveryPlans = @($RPRecId)
    $F.Cust_RiskAssess    = $RARecIds.ToArray()
    $F.Cust_SecPolicies   = $SecurityPolicyRecIDs
    $F.Cust_TrainVideos   = $TrainingVideoRecIDs
}
Update-Record -AppId $AppId.Customers -RecId $CustomerRecId -Fields $custPatch | Out-Null
Write-Host "   OK: Customer record updated"

# ── Summary ──────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "======================================================"
Write-Host "  PCS Onboarding Complete: $CustomerName"
Write-Host "======================================================"
Write-Host "  Customer RecID           : $CustomerRecId"
Write-Host "  CAM RecID                : $CAMRecId"
Write-Host "  CRB RecID                : $CRBRecId"
Write-Host "  Recovery Plans RecID     : $RPRecId"
Write-Host "  Risk Assessments created : $($RARecIds.Count)"
Write-Host "  Security Policies patched: $($SecurityPolicyRecIDs.Count)"
Write-Host "  Training Videos patched  : $($TrainingVideoRecIDs.Count)"
Write-Host "======================================================"
