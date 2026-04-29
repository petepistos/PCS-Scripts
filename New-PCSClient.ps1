<#
.SYNOPSIS
    Pistos Compliance Sentinel - New Client Onboarding Script (v2.0)

.DESCRIPTION
    Aligned with the Risk Assessment Template (Master sheet) as the source of truth.

    Control library:
      Process     : P-01 through P-13       (13 always-applicable, flat)
      Workstation : W-1.1 through W-1.5     (created if 'W' in -TechEnvironments)
      M365        : M-1.1 through M-1.5     (created if 'M' in -TechEnvironments)
      External    : E-1.1 through E-1.5     (created if 'E' in -TechEnvironments)
      Google WS   : G-1.1 through G-1.5     (created if 'G' in -TechEnvironments)
      Linux       : L-1.1 through L-1.5     (created if 'L' in -TechEnvironments)
      AMS         : AM-01 through AM-07     (created if -IncludeAMS, flat)

    Numbering: P and AM are flat (no parent aggregate); W/M/E/G/L use decimal
    notation because they are sub-controls of an environment aggregate that
    rolls up to a single environment score (e.g., W-1.1..W-1.5 -> W).

    Total records range: 18 (W only, no AMS) -> 45 (all environments + AMS).
    Default scope (W/M/E + AMS) yields 35 records.

    Execution order:
      1.  Look up Customer record by name, capture RecID
      2.  Create Customer Annual Metrics (CAM) record
      3.  Create Customer Receipts and Billing (CRB) record
      4.  Create Recovery and Inventory Plans record
      5.  Build in-scope control list and create Risk Assessment records linked to CAM
      6.  PATCH CAM with CRB link and all Risk Assessment links
      7.  PATCH all Security Policy records to add customer link
      8.  PATCH all Training and Awareness Video records to add customer link
      9.  PATCH Customer record with all links
     10.  Generate IRDR Plan portal token, push seed, save token

.EXAMPLE
    # Default scope: W/M/E + AMS (35 records)
    .\New-PCSClient.ps1 -CustomerName "Apex Brokerage" -ApiKey "your_api_key_here"

.EXAMPLE
    # Google Workspace shop, no Linux, with AMS
    .\New-PCSClient.ps1 -CustomerName "Acme Insurance" -ApiKey "..." -TechEnvironments W,G,E

.EXAMPLE
    # Process + AMS only, no technology scans yet
    .\New-PCSClient.ps1 -CustomerName "Smith Agency" -ApiKey "..." -TechEnvironments @()

.NOTES
    Run from C:\pistos\
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$CustomerName,
    [string]$AssessmentYear = "2025",
    [Parameter(Mandatory)] [string]$ApiKey,

    # Which technology environments to assess for this client.
    # Valid: W (Windows), M (M365), E (External Scan), G (Google Workspace), L (Linux)
    [ValidateSet('W','M','E','G','L')]
    [string[]]$TechEnvironments = @('W','M','E'),

    # Whether to create the AM-1.x AMS control records.
    [bool]$IncludeAMS = $true
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
# FIELD SLUGS
# =============================================================================
$F = @{
    Cust_Title          = "title"
    Cust_CAM            = "shatf2oq"
    Cust_CRB            = "sk5iewd5"
    Cust_RecoveryPlans  = "se5f1a15f6"
    Cust_RiskAssess     = "s40zjaz2"
    Cust_SecPolicies    = "sjvb15al"
    Cust_TrainVideos    = "s6a41a897e"
    Cust_SyncToken      = "s81dc03995"
    Cust_Phone          = "s5e03d44a2"
    Cust_Email          = "s000da1c94"
    CAM_Title           = "title"
    CAM_Customer        = "s66338ee26"
    CAM_CRB             = "sjq216p5"
    CAM_RiskAssess      = "s6n9jtk9"
    CAM_AssessYear      = "sa7da430dc"
    CRB_Title           = "title"
    CRB_Customer        = "sa87f7e561"
    CRB_AssessYear      = "s129d8489c"
    RP_Title            = "title"
    RP_Customer         = "s7bfc29230"
    RP_OCP_Template     = "sac8394f24"
    RP_User_Guide       = "sn3ojeg1"
    RP_IRP_DRP_Plan     = "s4310618ef"
    RP_IRP_DRP_Guide    = "s3c2k5bv"
    RA_Title            = "title"
    RA_ControlNumber    = "s6e4b09215"
    RA_ControlName      = "s2ca6be6cb"
    RA_InherentRisk     = "s45d89188b"
    RA_Applicable       = "scb215244e"
    RA_EvidenceStatus   = "sb4a4b9b71"
    RA_ImplGuide        = "s935240aaa"
    RA_Domain           = "s20jaedo"
    RA_AssessYear       = "s1c03507e0"
    RA_LinkCAM          = "s5a5c0cdc2"
    RA_LinkPolicies     = "sa0dd2faa7"
    RA_LinkNYCRR        = "s326dedc01"
    SP_Customers        = "s6335f6ddd"
    TV_Customers        = "sa4df46192"
}

# =============================================================================
# STATIC RecID ARRAYS
# =============================================================================
$SecurityPolicyRecIDs = @(
    "699d1e62306fc9aaa33057ba",  # 0  Asset Management
    "699c68284b668835aae4abcf",  # 1  Access Control & Identity Management
    "699c68284b668835aae4abd4",  # 2  Business Continuity & Disaster Recovery
    "699c70649c55ae0879ead3fc",  # 3  Customer & Consumer Privacy
    "699c68284b668835aae4abcd",  # 4  Cybersecurity Governance & Program
    "699c68284b668835aae4abd6",  # 5  Data Governance, Retention & Destruction
    "699c719f3bbbbb9c29edbbda",  # 6  Email Authentication and Anti-Spoofing
    "699c68284b668835aae4abd0",  # 7  Encryption & Data Protection
    "699c68284b668835aae4abd1",  # 8  Endpoint Security & Firewall
    "699c6c7f9c55ae0879ead3f9",  # 9  Human Resources
    "699c68284b668835aae4abd3",  # 10 Incident Response & Breach Notification
    "699c68284b668835aae4abd2",  # 11 Logging & Monitoring
    "699c68284b668835aae4abce",  # 12 Risk Assessment
    "699c68284b668835aae4abd5",  # 13 Vendor & Third-Party Risk Management
    "69c917978fe4bf681ad41990"   # 14 Environmental and Physical Security
)

$TrainingVideoRecIDs = @(
    "6931c3b826889575bdd457ee","6931c3b826889575bdd457ef","6931c3b826889575bdd457f0",
    "6931c3b826889575bdd457f1","6931c3b826889575bdd457f2","6931c3b826889575bdd457f3",
    "6931c3b826889575bdd457f4","6931c3b826889575bdd457f5","6931c3b826889575bdd457f6",
    "6931c3b826889575bdd457f7","6931c3b826889575bdd457f8","6931c3b826889575bdd457f9",
    "6931c3b826889575bdd457fa","6931c3b826889575bdd457fb","6931c3b826889575bdd457fc",
    "6931c3b826889575bdd457fd","6931c3b826889575bdd457fe","6931c3b826889575bdd457ff",
    "6931c3b826889575bdd45800","6931c3b826889575bdd45801","6931c3b826889575bdd45802",
    "6931c3b826889575bdd45803","6931c3b826889575bdd45804","6931c3b826889575bdd45805",
    "6931c7b40baefadbec68656a","6931c7b40baefadbec68656b","6931c7b40baefadbec68656c",
    "6931c7b40baefadbec68656d","6931c7b40baefadbec68656e","6931c7b40baefadbec68656f",
    "6931c7b40baefadbec686570","6931c7b40baefadbec686571","6931c7b40baefadbec686572",
    "6931c7b40baefadbec686573","6931c7b40baefadbec686574","6931c7b40baefadbec686575",
    "6931c7b40baefadbec686576","6931c7b40baefadbec686577","6931c7b40baefadbec686578",
    "6931c7b40baefadbec686579","6931c7b40baefadbec68657a","6931c7b40baefadbec68657b",
    "6931c7b40baefadbec68657c","6931c7b40baefadbec68657d","6931c7b40baefadbec68657e",
    "6931c7b40baefadbec68657f","6931c7b40baefadbec686580","6931c7b40baefadbec686581",
    "6931c7b40baefadbec686582","6931c7b40baefadbec686583","6931c7b40baefadbec686584",
    "6931c7b40baefadbec686585","6931c7b40baefadbec686586","6931c7b40baefadbec686587",
    "6931c7b40baefadbec686588","6931c7b40baefadbec686589","6931c7b40baefadbec68658a",
    "698df592f5cfce2b77ba8813","698df64d6b6914088ab96b33",
    "698e099b57c688b009efdec8","698e09fccc61f9c157c74c60"
)

# Policy Num -> RecID map
$PolicyRecIDMap = @{
    "0"="699d1e62306fc9aaa33057ba"; "1"="699c68284b668835aae4abcf"
    "2"="699c68284b668835aae4abd4"; "3"="699c70649c55ae0879ead3fc"
    "4"="699c68284b668835aae4abcd"; "5"="699c68284b668835aae4abd6"
    "6"="699c719f3bbbbb9c29edbbda"; "7"="699c68284b668835aae4abd0"
    "8"="699c68284b668835aae4abd1"; "9"="699c6c7f9c55ae0879ead3f9"
    "10"="699c68284b668835aae4abd3"; "11"="699c68284b668835aae4abd2"
    "12"="699c68284b668835aae4abce"; "13"="699c68284b668835aae4abd5"
    "14"="69c917978fe4bf681ad41990"
}
function Get-PolicyRecIDs { param([string[]]$PolicyNums)
    return $PolicyNums | ForEach-Object { $PolicyRecIDMap["$_"] } }

# NYCRR Section -> RecID map
$NYCRRRecIDMap = @{
    "500.01(c)"="69a64299a8c9231365e67c1a"; "500.01(e)"="69a64321a8c9231365e67c1b"
    "500.01(f)"="69a63d67d90a5c1409d0d9cc"; "500.01(j)"="69a63d6ad90a5c1409d0d9cd"
    "500.01(k)"="69a64507a8c9231365e67c27"; "500.01(l)"="69a64505a8c9231365e67c26"
    "500.01(n)p"="69a64405a8c9231365e67c1e"; "500.01(n)t"="69a63d81d90a5c1409d0d9d0"
    "500.01(p)"="69a63d82a028eeee017558cf"; "500.02"="699c8007ae8602aaab02bb5f"
    "500.03"="699c8007ae8602aaab02bb60"; "500.04"="699c8007ae8602aaab02bb61"
    "500.05"="699c8007ae8602aaab02bb63"; "500.06"="699c8007ae8602aaab02bb64"
    "500.07"="699c8007ae8602aaab02bb65"; "500.08"="699c8007ae8602aaab02bb66"
    "500.09"="699c8007ae8602aaab02bb67"; "500.10"="699c8007ae8602aaab02bb6a"
    "500.11"="699c8007ae8602aaab02bb6b"; "500.12"="699c8007ae8602aaab02bb6d"
    "500.13"="699c8007ae8602aaab02bb6e"; "500.14"="699c8007ae8602aaab02bb71"
    "500.15a"="699c8007ae8602aaab02bb73"; "500.15b"="699c8007ae8602aaab02bb74"
    "500.16"="699c8007ae8602aaab02bb75"; "500.17"="699c8007ae8602aaab02bb77"
    "500.18"="699c8007ae8602aaab02bb78"; "500.19"="69a64c1ed21bdee921801e68"
    "500.20"="69a64c0c58185fb14bba8a80"; "500.21"="69a64c1ed21bdee921801e67"
}
function Get-NYCRRRecIDs { param([string[]]$Sections)
    return $Sections | ForEach-Object { $NYCRRRecIDMap[$_] } }

# =============================================================================
# CONTROL LIBRARY (per Risk Assessment Template Master sheet, source of truth)
# =============================================================================
$REM_W   = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/Windows%20Workstation"
$REM_M   = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/MS%20365"
$REM_AM  = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/AMS"
$REM_E   = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/External%20Scans"
$REM_G   = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/Google%20Workspace"
$REM_L   = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/Linux"

# Process control SharePoint evidence-template links from spreadsheet K column.
# NOTE: these are under /pete/ in personal SharePoint. Re-share to tenant-level
# before client distribution.
$SP_P15_PrivacyNotice = "https://pistosip-my.sharepoint.com/:w:/p/pete/IQB14An3ZR_GTpWKpkTWphKHAbcuv4nwcgZNQ28BrNRgFSg?e=3OnNMm"
$SP_P16_VendorDD      = "https://pistosip-my.sharepoint.com/:f:/p/pete/IgCptSKX6tvYQLi4c7Xfje-0AYkyGRmaviq38dnpkvGnFKw?e=IPyU4q"
$SP_P17_Handbook      = "https://pistosip-my.sharepoint.com/:w:/p/pete/IQCs0-KfA1QyS6tibsprjE55AfKu-p2b48H_K0mbDuhKHBo?e=8h8cUE"
$SP_P19_Onboarding    = "https://pistosip-my.sharepoint.com/:w:/p/pete/IQD4P31eV9zfQqb6vdSf3g47AZmEW39mpCRzDTHmCCsH_-c?e=YSptoZ"
$SP_P110_RemoteAccess = "https://pistosip-my.sharepoint.com/:w:/p/pete/IQDRhGh_mDWGTq8TfDalngquAZWEEF2SF3c5ne-33dZGs6c?e=5nWPcn"
$SP_P112_CyberIns     = "https://pistosip-my.sharepoint.com/:w:/p/pete/IQBPfq5YXdbGQpCOx769Cno7AauikiTJ4G5-dzQzAVtYWpo?e=WOmWDP"
$SP_AM_EZLynxProc     = "https://pistosip-my.sharepoint.com/my?id=%2Fpersonal%2Fpete%5Fpistosip%5Fcom%2FDocuments%2FDesktop%2FPCS%2FControls%20Implementation%20Steps%2FAMS%2FEzLynx%2FEZLynx%5FNYDFS%5FCybersecurity%5FProcedure"

# ----- Process Controls (always created; flat numbering P-01..P-13) -----
$ProcessControls = @(
    @{ ControlNumber="P-01"; ControlName="Security policies addressing each applicable regulatory requirement have been documented, reviewed, and formally approved by management."; Domain="Governance & Program";          Applicable="Applicable"; EvidenceStatus="Accepted";              InherentRisk=10; Policies=@("4");      NYDFSRefs=@("500.03","500.04");           ImplGuide=$null },
    @{ ControlNumber="P-02"; ControlName="A documented Disaster Recovery Plan is maintained and tested on a defined schedule.";                                                       Domain="Operational Continuity";       Applicable="Applicable"; EvidenceStatus="Partially Implemented"; InherentRisk=5;  Policies=@("2");      NYDFSRefs=@("500.16");                    ImplGuide="https://petepistos.github.io/IRDR-Plan/" },
    @{ ControlNumber="P-03"; ControlName="An inventory of hardware and software assets used to store, process and maintain NPI is maintained.";                                       Domain="Asset Management";             Applicable="Applicable"; EvidenceStatus="Partially Implemented"; InherentRisk=5;  Policies=@("0");      NYDFSRefs=@("500.13");                    ImplGuide="https://petepistos.github.io/IRDR-Plan/" },
    @{ ControlNumber="P-04"; ControlName="A documented Incident Response Plan is maintained, including roles, escalation paths, and response procedures.";                            Domain="Operational Continuity";       Applicable="Applicable"; EvidenceStatus="Partially Implemented"; InherentRisk=10; Policies=@("10");     NYDFSRefs=@("500.16","500.17");           ImplGuide="https://petepistos.github.io/IRDR-Plan/" },
    @{ ControlNumber="P-05"; ControlName="A published privacy notice is maintained and includes an opt-out mechanism where applicable.";                                              Domain="Customer and Consumer Privacy";Applicable="Applicable"; EvidenceStatus="Partially Implemented"; InherentRisk=6;  Policies=@("3");      NYDFSRefs=@("500.18","500.01(k)");        ImplGuide=$SP_P15_PrivacyNotice },
    @{ ControlNumber="P-06"; ControlName="A vendor due diligence process is implemented, including annual reviews of third-party service providers.";                                 Domain="Risk Management";              Applicable="Applicable"; EvidenceStatus="Accepted";              InherentRisk=7;  Policies=@("13");     NYDFSRefs=@("500.11","500.01(n)t");       ImplGuide=$SP_P16_VendorDD },
    @{ ControlNumber="P-07"; ControlName="The agency maintains and annually updates a written cybersecurity employee handbook that defines acceptable use of carrier portals, the Agency Management System, and office automation tools, and requires every employee, contractor, and authorized user to acknowledge the handbook in writing prior to being granted access to Nonpublic Information and at least annually thereafter."; Domain="Human Resources"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=6; Policies=@("9");  NYDFSRefs=@("500.03","500.14"); ImplGuide=$SP_P17_Handbook },
    @{ ControlNumber="P-08"; ControlName="Phishing simulations are conducted at least monthly and results are tracked for remediation.";                                              Domain="Human Resources";              Applicable="Applicable"; EvidenceStatus="Accepted";              InherentRisk=10; Policies=@("9");      NYDFSRefs=@("500.14");                    ImplGuide=$null },
    @{ ControlNumber="P-09"; ControlName="Pre-employment background checks are completed for all hires as part of an Onboarding Checklist.";                                          Domain="Human Resources";              Applicable="Applicable"; EvidenceStatus="Partially Implemented"; InherentRisk=7;  Policies=@("9");      NYDFSRefs=@("500.10","500.14");           ImplGuide=$SP_P19_Onboarding },
    @{ ControlNumber="P-10"; ControlName="Remote access to internal systems requires an approved VPN with MFA.";                                                                       Domain="Identity and Access Management";Applicable="Applicable"; EvidenceStatus="Not Started";           InherentRisk=9;  Policies=@("1");      NYDFSRefs=@("500.07","500.12","500.15a"); ImplGuide=$SP_P110_RemoteAccess },
    @{ ControlNumber="P-11"; ControlName="Security awareness training is required and tailored to each user's role and technical responsibilities.";                                  Domain="Human Resources";              Applicable="Applicable"; EvidenceStatus="Accepted";              InherentRisk=8;  Policies=@("9");      NYDFSRefs=@("500.14");                    ImplGuide=$null },
    @{ ControlNumber="P-12"; ControlName="The agency maintains cyberinsurance coverage appropriate to its risk profile and contractual obligations.";                                 Domain="Operational Continuity";       Applicable="Applicable"; EvidenceStatus="Not Started";           InherentRisk=7;  Policies=@("4");      NYDFSRefs=@("500.02","500.04");           ImplGuide=$SP_P112_CyberIns },
    @{ ControlNumber="P-13"; ControlName="The agency performs annual risk assessments that assess compliance with the NY DFS Cybersecurity regulation.";                              Domain="Risk Management";              Applicable="Applicable"; EvidenceStatus="Partially Implemented"; InherentRisk=10; Policies=@("12");     NYDFSRefs=@("500.09","500.01(p)");        ImplGuide=$null }
)

# ----- Workstation Controls (W) -----
$WControls = @(
    @{ ControlNumber="W-1.1"; ControlName="Identity & Access Control: Account Lockout, User Inventory, Password Policy, Admin Group Membership, Standard User Admin Rights, Break-Glass Account, UAC Level, User Rights Assignments";       Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=6; Policies=@("1","8");  NYDFSRefs=@("500.07","500.12","500.01(n)p"); ImplGuide=$REM_W },
    @{ ControlNumber="W-1.2"; ControlName="Data Protection: BitLocker Encryption, Secure Boot, Cached Domain Credentials, Credential Guard / LSA Protection";                                                                                Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=4; Policies=@("7","8");  NYDFSRefs=@("500.15a","500.15b");            ImplGuide=$REM_W },
    @{ ControlNumber="W-1.3"; ControlName="Threat Defense & Hardening: Software Inventory, AutoRun / AutoPlay, PowerShell Execution Policy, Windows Script Host, LLMNR / NetBIOS, WinRM, Remote Desktop / NLA, Windows Update, Defender Firewall, Local Admin Audit"; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=8; Policies=@("0","8");  NYDFSRefs=@("500.05","500.08","500.13");     ImplGuide=$REM_W },
    @{ ControlNumber="W-1.4"; ControlName="Logging & Monitoring: Audit Policy, Event Log Sizes, Sysmon, PowerShell Logging, Endpoint Protection Monitoring";                                                                                  Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=5; Policies=@("11");     NYDFSRefs=@("500.06","500.14");              ImplGuide=$REM_W },
    @{ ControlNumber="W-1.5"; ControlName="External Posture & Email: Secure DNS Configuration";                                                                                                                                              Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("8","6");  NYDFSRefs=@("500.08");                       ImplGuide=$REM_W }
)

# ----- M365 Controls (M) -----
$MControls = @(
    @{ ControlNumber="M-1.1"; ControlName="Identity & Access Control: Smart Lockout, User Account Inventory, Multifactor Authentication, Password Policy, Administrative Privilege Review, Conditional Access, Guest Access Restrictions, Legacy Authentication"; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=8; Policies=@("1");      NYDFSRefs=@("500.07","500.12","500.01(n)p"); ImplGuide=$REM_M },
    @{ ControlNumber="M-1.2"; ControlName="Data Protection: Intune Device Encryption Compliance, Data Loss Prevention Policies, SharePoint / OneDrive External Sharing";                                                                       Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=5; Policies=@("5","7");  NYDFSRefs=@("500.13","500.15a","500.01(k)"); ImplGuide=$REM_M },
    @{ ControlNumber="M-1.3"; ControlName="Threat Defense & Hardening: Safe Attachments / Safe Links, Anti-Phishing Policy";                                                                                                                  Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=3; Policies=@("8","6");  NYDFSRefs=@("500.08","500.14");              ImplGuide=$REM_M },
    @{ ControlNumber="M-1.4"; ControlName="Logging & Monitoring: Unified Audit Log, Audit Log Retention, Microsoft Secure Score";                                                                                                             Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=5; Policies=@("11");     NYDFSRefs=@("500.06","500.14");              ImplGuide=$REM_M },
    @{ ControlNumber="M-1.5"; ControlName="External Posture & Email: Mail Forwarding Controls, Shared Mailbox Authentication";                                                                                                                Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=4; Policies=@("1","6");  NYDFSRefs=@("500.07","500.08");              ImplGuide=$REM_M }
)

# ----- External Scan Controls (E) -----
$EControls = @(
    @{ ControlNumber="E-1.1"; ControlName="Identity & Access Control: (not applicable to external scan)";                                                                                                                                    Domain="Technology"; Applicable="N/A";        EvidenceStatus="Not Started"; InherentRisk=0;  Policies=@();         NYDFSRefs=@();                            ImplGuide=$REM_E },
    @{ ControlNumber="E-1.2"; ControlName="Data Protection: TLS Certificate Validation, TLS Version Enforcement";                                                                                                                            Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=3;  Policies=@("7","8");  NYDFSRefs=@("500.15a","500.15b");         ImplGuide=$REM_E },
    @{ ControlNumber="E-1.3"; ControlName="Threat Defense & Hardening: Port Scan, HTTP Methods, Open Redirect, Directory Listing, Exposed Admin Paths, HTTP Protocol Version";                                                               Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=5;  Policies=@("8","4");  NYDFSRefs=@("500.08");                    ImplGuide=$REM_E },
    @{ ControlNumber="E-1.4"; ControlName="Logging & Monitoring: (not applicable to external scan)";                                                                                                                                         Domain="Technology"; Applicable="N/A";        EvidenceStatus="Not Started"; InherentRisk=0;  Policies=@();         NYDFSRefs=@();                            ImplGuide=$REM_E },
    @{ ControlNumber="E-1.5"; ControlName="External Posture & Email: Security Headers, Permissions-Policy, CORS Configuration, SPF Record, DMARC Policy, DKIM Signing, CAA Record, DNSSEC, Zone Transfer Protection, MX Records, MTA-STS / TLSRPT, Email Authentication"; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=17; Policies=@("6","4","8"); NYDFSRefs=@("500.08","500.15a"); ImplGuide=$REM_E }
)

# ----- Google Workspace Controls (G) -----
$GControls = @(
    @{ ControlNumber="G-1.1"; ControlName="Identity & Access Control: Account Lockout, User Account Inventory, Multifactor Authentication, Password Policy, Administrative Privilege Review, Context-Aware Access, Guest/External Access Restrictions, Legacy Protocol Enforcement"; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=8; Policies=@("1");     NYDFSRefs=@("500.07","500.12","500.01(n)p"); ImplGuide=$REM_G },
    @{ ControlNumber="G-1.2"; ControlName="Data Protection: Drive Encryption, Data Loss Prevention Rules, Drive External Sharing Restrictions";                                                                                              Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=5; Policies=@("5","7"); NYDFSRefs=@("500.13","500.15a","500.01(k)"); ImplGuide=$REM_G },
    @{ ControlNumber="G-1.3"; ControlName="Threat Defense & Hardening: Gmail Advanced Protection, Anti-Phishing and Anti-Malware Settings";                                                                                                  Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=3; Policies=@("8","6"); NYDFSRefs=@("500.08","500.14");             ImplGuide=$REM_G },
    @{ ControlNumber="G-1.4"; ControlName="Logging & Monitoring: Admin Audit Log, Login Audit Log, Drive Audit Log, Log Retention";                                                                                                          Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=5; Policies=@("11");    NYDFSRefs=@("500.06","500.14");             ImplGuide=$REM_G },
    @{ ControlNumber="G-1.5"; ControlName="External Posture & Email: Mail Forwarding Controls, Shared Inbox Authentication";                                                                                                                 Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=4; Policies=@("1","6"); NYDFSRefs=@("500.07","500.08");             ImplGuide=$REM_G }
)

# ----- Linux Endpoint Controls (L) -----
$LControls = @(
    @{ ControlNumber="L-1.1"; ControlName="Identity & Access Control: SSH Root Login, SSH Password Auth, Password Complexity, Account Lockout, Passwordless Sudo, Inactive Accounts, Legacy Trust Files, Home Directory Permissions";        Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=6;  Policies=@("1","8"); NYDFSRefs=@("500.07","500.12","500.01(n)p"); ImplGuide=$REM_L },
    @{ ControlNumber="L-1.2"; ControlName="Data Protection: LUKS Disk Encryption, AIDE File Integrity, GRUB Bootloader Password";                                                                                                            Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=3;  Policies=@("7","8"); NYDFSRefs=@("500.15a","500.15b");            ImplGuide=$REM_L },
    @{ ControlNumber="L-1.3"; ControlName="Threat Defense & Hardening: /tmp Mount Options, World-Writable Files, SUID/SGID Audit, Firewall Enabled, SSH Idle Timeout, Listening Ports, ICMP Redirects, IP Forwarding, Unattended Upgrades, Pending Updates, Kernel Hardening"; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=11; Policies=@("0","8"); NYDFSRefs=@("500.05","500.08","500.13");     ImplGuide=$REM_L },
    @{ ControlNumber="L-1.4"; ControlName="Logging & Monitoring: auditd, System Logging, Log Rotation, Failed Login Logging";                                                                                                                Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=5;  Policies=@("11");    NYDFSRefs=@("500.06","500.14");              ImplGuide=$REM_L },
    @{ ControlNumber="L-1.5"; ControlName="External Posture & Email: (not applicable to internal Linux endpoint)";                                                                                                                           Domain="Technology"; Applicable="N/A";        EvidenceStatus="Not Started"; InherentRisk=0;  Policies=@();         NYDFSRefs=@();                            ImplGuide=$REM_L }
)

# ----- AMS Controls (AM, flat numbering AM-01..AM-07) -----
# Default Applicable = "N/A equivalent technology" per Pete's design: AMS controls
# are presumed covered by W or M unless the AMS uniquely owns them. Reviewer flips
# Applicable to "Applicable" on a per-control basis if the AMS owns the control.
$AMControls = @(
    @{ ControlNumber="AM-01"; ControlName="Accounts lock out for at least 10 minutes after no fewer than five failed logon attempts.";                Domain="Identity and Access Management"; Applicable="N/A equivalent technology"; EvidenceStatus="Not Started"; InherentRisk=3; Policies=@("1");     NYDFSRefs=@("500.07","500.12");        ImplGuide=$SP_AM_EZLynxProc },
    @{ ControlNumber="AM-02"; ControlName="User accounts are uniquely assigned to and attributable to a single individual.";                          Domain="Identity and Access Management"; Applicable="N/A equivalent technology"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("1");     NYDFSRefs=@("500.07");                 ImplGuide=$SP_AM_EZLynxProc },
    @{ ControlNumber="AM-03"; ControlName="Logging and monitoring are enabled and configured to record security-relevant events.";                    Domain="Logging and Monitoring";         Applicable="N/A equivalent technology"; EvidenceStatus="Not Started"; InherentRisk=3; Policies=@("11");    NYDFSRefs=@("500.06","500.14");        ImplGuide=$SP_AM_EZLynxProc },
    @{ ControlNumber="AM-04"; ControlName="Passwords meet the organization's approved complexity and length requirements.";                           Domain="Identity and Access Management"; Applicable="N/A equivalent technology"; EvidenceStatus="Not Started"; InherentRisk=7; Policies=@("1");     NYDFSRefs=@("500.07","500.12");        ImplGuide=$SP_AM_EZLynxProc },
    @{ ControlNumber="AM-05"; ControlName="Roles and permissions are assigned to enforce least privilege and separation of duties.";                  Domain="Identity and Access Management"; Applicable="N/A equivalent technology"; EvidenceStatus="Not Started"; InherentRisk=3; Policies=@("1");     NYDFSRefs=@("500.07","500.01(n)p");    ImplGuide=$SP_AM_EZLynxProc },
    @{ ControlNumber="AM-06"; ControlName="The system uses 2FA where available.";                                                                     Domain="Identity and Access Management"; Applicable="N/A equivalent technology"; EvidenceStatus="Not Started"; InherentRisk=5; Policies=@("1");     NYDFSRefs=@("500.12");                 ImplGuide=$SP_AM_EZLynxProc },
    @{ ControlNumber="AM-07"; ControlName="Users do not have administrative privileges.";                                                             Domain="Identity and Access Management"; Applicable="N/A equivalent technology"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("1");     NYDFSRefs=@("500.07","500.01(n)p");    ImplGuide=$SP_AM_EZLynxProc }
)

# ----- Build in-scope control list -----
$Controls = [System.Collections.Generic.List[hashtable]]::new()
$ProcessControls | ForEach-Object { $Controls.Add($_) }
if ($TechEnvironments -contains 'W') { $WControls | ForEach-Object { $Controls.Add($_) } }
if ($TechEnvironments -contains 'M') { $MControls | ForEach-Object { $Controls.Add($_) } }
if ($TechEnvironments -contains 'E') { $EControls | ForEach-Object { $Controls.Add($_) } }
if ($TechEnvironments -contains 'G') { $GControls | ForEach-Object { $Controls.Add($_) } }
if ($TechEnvironments -contains 'L') { $LControls | ForEach-Object { $Controls.Add($_) } }
if ($IncludeAMS)                     { $AMControls | ForEach-Object { $Controls.Add($_) } }

# =============================================================================
# HELPERS
# =============================================================================
$Headers = @{
    "Authorization" = "Token $ApiKey"
    "Account-Id"    = $AccountId
    "Content-Type"  = "application/json"
}

function Invoke-SS {
    param([string]$Method, [string]$Url, [hashtable]$Body = $null)
    $params = @{ Method=$Method; Uri=$Url; Headers=$Headers }
    if ($Body) { $params.Body = ($Body | ConvertTo-Json -Depth 10 -Compress) }
    $result = Invoke-RestMethod @params
    Start-Sleep -Milliseconds 500
    return $result
}

function New-Record { param([string]$AppId, [hashtable]$Fields)
    return Invoke-SS -Method POST -Url "$BaseUrl/$AppId/records/" -Body $Fields }

function Update-Record { param([string]$AppId, [string]$RecId, [hashtable]$Fields)
    return Invoke-SS -Method PATCH -Url "$BaseUrl/$AppId/records/$RecId/" -Body $Fields }

function Get-Record { param([string]$AppId, [string]$RecId)
    return Invoke-SS -Method GET -Url "$BaseUrl/$AppId/records/$RecId/" }

function Find-CustomerRecId { param([string]$Name)
    $resp = Invoke-SS -Method POST -Url "$BaseUrl/$($AppId.Customers)/records/list/" -Body @{ filter=@{}; sort=@(); fields_to_search=@() }
    $match = $resp.items | Where-Object { $_.title -eq $Name }
    if (-not $match) { throw "Customer '$Name' not found in SmartSuite." }
    return $match.id }

# =============================================================================
# MAIN
# =============================================================================
Write-Host ""
Write-Host "======================================================"
Write-Host "  PCS Onboarding: $CustomerName"
Write-Host "  Tech Environments : $($TechEnvironments -join ', ')"
Write-Host "  Include AMS       : $IncludeAMS"
Write-Host "  Total Controls    : $($Controls.Count)"
Write-Host "======================================================"

Write-Host ""; Write-Host ">> Looking up Customer: $CustomerName"
$CustomerRecId = Find-CustomerRecId -Name $CustomerName
Write-Host "   OK: Customer RecID: $CustomerRecId"

Write-Host ""; Write-Host ">> Creating Customer Annual Metrics record"
$camResp  = New-Record -AppId $AppId.CAM -Fields @{ $F.CAM_Title=$CustomerName; $F.CAM_Customer=@($CustomerRecId); $F.CAM_AssessYear=$AssessmentYear }
$CAMRecId = $camResp.id
Write-Host "   OK: CAM RecID: $CAMRecId"

Write-Host ""; Write-Host ">> Creating Customer Receipts and Billing record"
$crbResp  = New-Record -AppId $AppId.CRB -Fields @{ $F.CRB_Title=$CustomerName; $F.CRB_Customer=@($CustomerRecId); $F.CRB_AssessYear=$AssessmentYear }
$CRBRecId = $crbResp.id
Write-Host "   OK: CRB RecID: $CRBRecId"

Write-Host ""; Write-Host ">> Creating Recovery and Inventory Plans record"
$rpResp  = New-Record -AppId $AppId.RecoveryPlans -Fields @{
    $F.RP_Title=$CustomerName; $F.RP_Customer=@($CustomerRecId)
    $F.RP_OCP_Template="https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Recovery-Plans/IRP_DRP_Plan_with_User_Guide.xlsx"
    $F.RP_User_Guide="https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Recovery-Plans/IRP_DRP_User_Guide.pdf"
    $F.RP_IRP_DRP_Plan="https://petepistos.github.io/skopein-scripts/Recovery-Plans/IRP_DRP_Plan.html"
    $F.RP_IRP_DRP_Guide="https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Recovery-Plans/IRP_DRP_User_Guide.pdf"
}
$RPRecId = $rpResp.id
Write-Host "   OK: Recovery Plans RecID: $RPRecId"

Write-Host ""; Write-Host ">> Creating $($Controls.Count) Risk Assessment records"
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
    if ($null -ne $ctrl.ImplGuide) { $raBody[$F.RA_ImplGuide] = $ctrl.ImplGuide }
    if ($ctrl.Policies -and $ctrl.Policies.Count -gt 0) {
        $raBody[$F.RA_LinkPolicies] = @(Get-PolicyRecIDs -PolicyNums $ctrl.Policies)
    }
    if ($ctrl.NYDFSRefs -and $ctrl.NYDFSRefs.Count -gt 0) {
        $raBody[$F.RA_LinkNYCRR] = @(Get-NYCRRRecIDs -Sections $ctrl.NYDFSRefs)
    }
    $raResp = New-Record -AppId $AppId.RiskAssessments -Fields $raBody
    $RARecIds.Add($raResp.id)
    Write-Host "   OK: $($ctrl.ControlNumber) created"
}
Write-Host "   OK: $($Controls.Count) of $($Controls.Count) Risk Assessment records created"

Write-Host ""; Write-Host ">> Updating CAM with CRB and Risk Assessment links"
Update-Record -AppId $AppId.CAM -RecId $CAMRecId -Fields @{ $F.CAM_CRB=@($CRBRecId); $F.CAM_RiskAssess=$RARecIds.ToArray() } | Out-Null
Write-Host "   OK: CAM updated"

Write-Host ""; Write-Host ">> Adding customer to all $($SecurityPolicyRecIDs.Count) Security Policy records"
foreach ($spId in $SecurityPolicyRecIDs) {
    Update-Record -AppId $AppId.SecurityPolicies -RecId $spId -Fields @{ $F.SP_Customers=@($CustomerRecId) } | Out-Null
    Write-Host "   OK: Policy $spId updated"
}

Write-Host ""; Write-Host ">> Adding customer to all $($TrainingVideoRecIDs.Count) Training and Awareness Video records"
foreach ($tvId in $TrainingVideoRecIDs) {
    Update-Record -AppId $AppId.TrainingVideos -RecId $tvId -Fields @{ $F.TV_Customers=@($CustomerRecId) } | Out-Null
    Write-Host "   OK: Video $tvId updated"
}

Write-Host ""; Write-Host ">> Updating Customer record with all links"
Update-Record -AppId $AppId.Customers -RecId $CustomerRecId -Fields @{
    $F.Cust_CAM=@($CAMRecId); $F.Cust_CRB=@($CRBRecId)
    $F.Cust_RecoveryPlans=@($RPRecId); $F.Cust_RiskAssess=$RARecIds.ToArray()
    $F.Cust_SecPolicies=$SecurityPolicyRecIDs; $F.Cust_TrainVideos=$TrainingVideoRecIDs
} | Out-Null
Write-Host "   OK: Customer record updated"

# =============================================================================
# IRDR PLAN PORTAL - generate token, push seed, save token to Customer
# =============================================================================
Write-Host ""; Write-Host ">> Setting up IRDR Plan portal access"

$IrdrToken = [guid]::NewGuid().ToString('N')

$CustomerRecord = Get-Record -AppId $AppId.Customers -RecId $CustomerRecId
$ContactPhone = $CustomerRecord.($F.Cust_Phone)
$ContactEmail = $CustomerRecord.($F.Cust_Email)

$SeedBody = @{
    seedToken = $IrdrToken
    seedData = @{
        profile = @{
            companyLegalName = $CustomerName
            mainPhone        = $ContactPhone
            generalEmail     = $ContactEmail
        }
    }
} | ConvertTo-Json -Depth 5 -Compress

try {
    $SeedResp = Invoke-RestMethod `
        -Uri 'https://eo4oxn8yubmyf3d.m.pipedream.net' `
        -Method Post -ContentType 'application/json' -Body $SeedBody
    if ($SeedResp.ok) { Write-Host "   OK: Seed pushed to portal" }
    else { Write-Host "   WARN: Unexpected seed response: $($SeedResp | ConvertTo-Json -Compress)" -ForegroundColor Yellow }
} catch {
    Write-Host "   WARN: Seed push failed: $_" -ForegroundColor Yellow
    Write-Host "         Token will still be saved on Customer record" -ForegroundColor Yellow
}

Update-Record -AppId $AppId.Customers -RecId $CustomerRecId -Fields @{
    $F.Cust_SyncToken = $IrdrToken
} | Out-Null
Write-Host "   OK: Token saved to Customer record"

$ClientIrdrUrl = "https://petepistos.github.io/IRDR-Plan/?c=$IrdrToken&seed=1"

# =============================================================================
# SUMMARY
# =============================================================================
$TotalInherentRisk = ($Controls | ForEach-Object { $_.InherentRisk } | Measure-Object -Sum).Sum

Write-Host ""
Write-Host "======================================================"
Write-Host "  PCS Onboarding Complete: $CustomerName"
Write-Host "======================================================"
Write-Host "  Customer RecID            : $CustomerRecId"
Write-Host "  CAM RecID                 : $CAMRecId"
Write-Host "  CRB RecID                 : $CRBRecId"
Write-Host "  Recovery Plans RecID      : $RPRecId"
Write-Host "  Tech Environments         : $($TechEnvironments -join ', ')"
Write-Host "  Include AMS               : $IncludeAMS"
Write-Host "  Risk Assessments created  : $($RARecIds.Count)"
Write-Host "    Process (P-01..P-13)    : 13"
if ($TechEnvironments -contains 'W') { Write-Host "    Workstation (W-1.x)     : 5" }
if ($TechEnvironments -contains 'M') { Write-Host "    M365 (M-1.x)            : 5" }
if ($TechEnvironments -contains 'E') { Write-Host "    External (E-1.x)        : 5" }
if ($TechEnvironments -contains 'G') { Write-Host "    Google WS (G-1.x)       : 5" }
if ($TechEnvironments -contains 'L') { Write-Host "    Linux (L-1.x)           : 5" }
if ($IncludeAMS)                     { Write-Host "    AMS (AM-01..AM-07)      : 7" }
Write-Host "  Total Inherent Risk Pts   : $TotalInherentRisk"
Write-Host "  Security Policies patched : $($SecurityPolicyRecIDs.Count)"
Write-Host "  Training Videos patched   : $($TrainingVideoRecIDs.Count)"
Write-Host "  IRDR Sync Token           : $IrdrToken"
Write-Host "  IRDR Client URL           : $ClientIrdrUrl"
Write-Host "======================================================"
