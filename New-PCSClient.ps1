```powershell
<#
.SYNOPSIS
    Pistos Compliance Sentinel - New Client Onboarding Script

.DESCRIPTION
    Execution order:
      1.  Look up Customer record by name, capture RecID
      2.  Create Customer Annual Metrics (CAM) record
      3.  Create Customer Receipts and Billing (CRB) record
      4.  Create Recovery and Inventory Plans record
      5.  Create 55 Risk Assessment records linked to CAM
      6.  PATCH CAM with CRB link and all 55 Risk Assessment links
      7.  PATCH all 15 Security Policy records to add customer link
      8.  PATCH all 61 Training and Awareness Video records to add customer link
      9.  PATCH Customer record with all links

.EXAMPLE
    .\New-PCSClient.ps1 -CustomerName "Apex Brokerage" -ApiKey "your_api_key_here"

.NOTES
    Run from C:\pistos\
    Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory)] [string]$CustomerName,
    [string]$AssessmentYear = "2025",
    [Parameter(Mandatory)] [string]$ApiKey
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
    RA_Platform         = "s6bd84614f"
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

# Policy Num -> RecID map (string keys)
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
# CONTROL LIBRARY (55 controls)
# =============================================================================
$EFT     = "https://github.com/petepistos/skopein-scripts/tree/main/Evidence%20Form%20Templates"
$REM_W   = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/Windows%20Workstation"
$REM_M   = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/MS%20365"
$REM_AMS = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/AMS"
$REM_EXT = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/External%20Scans"
$REM_WA  = "https://github.com/petepistos/skopein-scripts/tree/main/Remediation_docs/WebApp"

function MakeUrl($url, $label) { if ([string]::IsNullOrEmpty($url)) { return $null }; return $url }

$Controls = @(
    # P Controls
    @{ ControlNumber="P-01"; ControlName="A documented Disaster Recovery Plan is maintained and tested on a defined schedule."; Domain="Operational Continuity"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=3; Policies=@("2"); NYDFSRefs=@("500.16"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-01_Disaster_Recovery_Plan_Reference.pdf" "P-01" },
    @{ ControlNumber="P-02"; ControlName="A documented Incident Response Plan is maintained, including roles, escalation paths, and response procedures."; Domain="Operational Continuity"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=5; Policies=@("10"); NYDFSRefs=@("500.16","500.17"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-02_Incident_Response_Plan_Reference.pdf" "P-02" },
    @{ ControlNumber="P-03"; ControlName="A formally documented patch management process is maintained, with defined SLAs based on severity."; Domain="Asset Management"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=3; Policies=@("0","8"); NYDFSRefs=@("500.05","500.13"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-03_Patch_Management_Reference.pdf" "P-03" },
    @{ ControlNumber="P-04"; ControlName="A published privacy notice is maintained and includes an opt-out mechanism where applicable."; Domain="Customer and Consumer Privacy"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=1; Policies=@("3"); NYDFSRefs=@("500.18","500.01(k)"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-04_Privacy_Notice_OptOut.pdf" "P-04" },
    @{ ControlNumber="P-05"; ControlName="A vendor due diligence process is implemented, including annual reviews of third-party service providers."; Domain="Risk Management"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=1; Policies=@("13"); NYDFSRefs=@("500.11","500.01(n)t"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-05_Vendor_Due_Diligence_Reference.pdf" "P-05" },
    @{ ControlNumber="P-06"; ControlName="Employees are responsible for securing laptops and mobile devices when unattended and lock screens when stepping away."; Domain="Human Resources"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=2; Policies=@("8"); NYDFSRefs=@("500.07","500.14"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-06_Device_ScreenLock_Acknowledgment.pdf" "P-06" },
    @{ ControlNumber="P-07"; ControlName="Employee onboarding and offboarding processes are documented."; Domain="Human Resources"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=1; Policies=@("9"); NYDFSRefs=@("500.07","500.14"); ImplGuide=MakeUrl $EFT "P-07" },
    @{ ControlNumber="P-08"; ControlName="Only authorized personnel can access company facilities based on defined job roles and business need."; Domain="Environmental Security"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=1; Policies=@("14"); NYDFSRefs=@("500.02","500.07"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-08_Physical_Access_Controls_Reference.pdf" "P-08" },
    @{ ControlNumber="P-09"; ControlName="Pre-employment background checks are completed for all hires, consistent with legal requirements."; Domain="Human Resources"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=2; Policies=@("9"); NYDFSRefs=@("500.10","500.14"); ImplGuide=MakeUrl $EFT "P-09" },
    @{ ControlNumber="P-10"; ControlName="Remote access to internal systems requires an approved VPN with MFA."; Domain="Identity and Access Management"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=5; Policies=@("1"); NYDFSRefs=@("500.07","500.12"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-10_Remote_Access_Request_Form.pdf" "P-10" },
    @{ ControlNumber="P-12"; ControlName="Secure configuration standards are documented and applied consistently to systems and endpoints."; Domain="Identity and Access Management"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=1; Policies=@("0","8"); NYDFSRefs=@("500.05","500.08","500.13"); ImplGuide=MakeUrl $EFT "P-12" },
    @{ ControlNumber="P-13"; ControlName="Security awareness training is required and tailored to each user's role and technical responsibilities."; Domain="Human Resources"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=4; Policies=@("9"); NYDFSRefs=@("500.14"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-13_Security_Awareness_Training_Reference.pdf" "P-13" },
    @{ ControlNumber="P-15"; ControlName="Security policies addressing each applicable regulatory requirement have been documented, reviewed, and formally approved by management."; Domain="Governance & Program"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=5; Policies=@("4"); NYDFSRefs=@("500.03","500.04"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-15_Security_Policy_Documentation_Reference.pdf" "P-15" },
    @{ ControlNumber="P-16"; ControlName="The entity is paperless."; Domain="Customer and Consumer Privacy"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=3; Policies=@("5"); NYDFSRefs=@("500.13","500.15b"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-16_Paperless_Office_Reference.pdf" "P-16" },
    @{ ControlNumber="P-17"; ControlName="The entity maintains cyberinsurance coverage appropriate to its risk profile and contractual obligations."; Domain="Operational Continuity"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("4"); NYDFSRefs=@("500.02","500.04"); ImplGuide=$null },
    @{ ControlNumber="P-18"; ControlName="The entity's risk management program performs annual risk assessments that assess compliance with the NY DFS Cybersecurity regulation."; Domain="Risk Management"; Applicable="Applicable"; EvidenceStatus="Accepted"; InherentRisk=5; Policies=@("12"); NYDFSRefs=@("500.09","500.01(p)"); ImplGuide=MakeUrl "https://raw.githubusercontent.com/petepistos/skopein-scripts/main/Evidence%20Form%20Templates/P-18_Annual_Risk_Assessment_Reference.pdf" "P-18" },
    @{ ControlNumber="P-19"; ControlName="Regular phishing simulations are conducted."; Domain="Human Resources"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=6; Policies=@("9"); NYDFSRefs=@("500.14"); ImplGuide=$null },

    # W Controls
    @{ ControlNumber="W-01"; ControlName="Endpoint protection is deployed and actively maintained on all workstations."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=3; Policies=@("8"); NYDFSRefs=@("500.05","500.08"); ImplGuide=MakeUrl $REM_W "W-01" },
    @{ ControlNumber="W-02"; ControlName="Full disk encryption is enabled on all workstations."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("7","8"); NYDFSRefs=@("500.15a","500.15b"); ImplGuide=MakeUrl $REM_W "W-02" },
    @{ ControlNumber="W-03"; ControlName="Operating system patches are applied within defined SLA timeframes based on severity."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=3; Policies=@("8","0"); NYDFSRefs=@("500.05","500.13"); ImplGuide=MakeUrl $REM_W "W-03" },
    @{ ControlNumber="W-04"; ControlName="Local administrator accounts are disabled or restricted on all workstations."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("1","8"); NYDFSRefs=@("500.07","500.01(n)p"); ImplGuide=MakeUrl $REM_W "W-04" },
    @{ ControlNumber="W-05"; ControlName="Screen lock and idle timeout policies are enforced on all workstations."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("8"); NYDFSRefs=@("500.07"); ImplGuide=MakeUrl $REM_W "W-05" },
    @{ ControlNumber="W-06"; ControlName="Host-based firewall is enabled and configured on all workstations."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("8"); NYDFSRefs=@("500.08"); ImplGuide=MakeUrl $REM_W "W-06" },
    @{ ControlNumber="W-07"; ControlName="Removable media usage is restricted or controlled by policy and technical enforcement."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("8"); NYDFSRefs=@("500.13","500.15a"); ImplGuide=MakeUrl $REM_W "W-07" },
    @{ ControlNumber="W-08"; ControlName="Only approved software is permitted to execute on workstations."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("8","0"); NYDFSRefs=@("500.08","500.13"); ImplGuide=MakeUrl $REM_W "W-08" },
    @{ ControlNumber="W-09"; ControlName="Audit logging is enabled and retained on all workstations."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("11"); NYDFSRefs=@("500.06","500.14"); ImplGuide=MakeUrl $REM_W "W-09" },
    @{ ControlNumber="W-10"; ControlName="Secure boot and BIOS/UEFI password protections are configured on all workstations."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("8"); NYDFSRefs=@("500.08"); ImplGuide=MakeUrl $REM_W "W-10" },
    @{ ControlNumber="W-11"; ControlName="Workstations are enrolled in a centralized device management platform."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("0"); NYDFSRefs=@("500.13"); ImplGuide=MakeUrl $REM_W "W-11" },

    # M Controls
    @{ ControlNumber="M-01"; ControlName="MFA is enforced for all Microsoft 365 user accounts."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=3; Policies=@("1"); NYDFSRefs=@("500.12"); ImplGuide=MakeUrl $REM_M "M-01" },
    @{ ControlNumber="M-02"; ControlName="Conditional Access policies are configured to restrict access based on risk and location."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=3; Policies=@("1"); NYDFSRefs=@("500.07","500.12"); ImplGuide=MakeUrl $REM_M "M-02" },
    @{ ControlNumber="M-03"; ControlName="Privileged administrative roles are separated from standard user accounts."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("1"); NYDFSRefs=@("500.07","500.01(n)p"); ImplGuide=MakeUrl $REM_M "M-03" },
    @{ ControlNumber="M-04"; ControlName="Mailbox audit logging is enabled for all Microsoft 365 users."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("11"); NYDFSRefs=@("500.06","500.14"); ImplGuide=MakeUrl $REM_M "M-04" },
    @{ ControlNumber="M-05"; ControlName="Anti-phishing and anti-spam policies are configured and enforced in Microsoft 365."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("8","6"); NYDFSRefs=@("500.08","500.14"); ImplGuide=MakeUrl $REM_M "M-05" },
    @{ ControlNumber="M-06"; ControlName="Safe Links and Safe Attachments (Defender for Office 365) are enabled for all users."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("8"); NYDFSRefs=@("500.08"); ImplGuide=MakeUrl $REM_M "M-06" },
    @{ ControlNumber="M-07"; ControlName="Data Loss Prevention (DLP) policies are defined and actively enforced."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("5"); NYDFSRefs=@("500.13","500.15a","500.01(k)"); ImplGuide=MakeUrl $REM_M "M-07" },
    @{ ControlNumber="M-08"; ControlName="External sharing for SharePoint and OneDrive is restricted to authorized use cases."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("1"); NYDFSRefs=@("500.07","500.13","500.15a"); ImplGuide=MakeUrl $REM_M "M-08" },
    @{ ControlNumber="M-09"; ControlName="The Unified Audit Log is enabled and retained per policy."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("11"); NYDFSRefs=@("500.06","500.14"); ImplGuide=MakeUrl $REM_M "M-09" },
    @{ ControlNumber="M-10"; ControlName="Legacy authentication protocols are blocked for all Microsoft 365 accounts."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("1"); NYDFSRefs=@("500.07","500.12"); ImplGuide=MakeUrl $REM_M "M-10" },
    @{ ControlNumber="M-11"; ControlName="Identity Protection risky sign-in alerts are configured and actively monitored."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("1"); NYDFSRefs=@("500.05","500.14"); ImplGuide=MakeUrl $REM_M "M-11" },

    # AMS Controls
    @{ ControlNumber="AMS-01"; ControlName="Smart Lockout is configured to prevent brute-force attacks on user accounts."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("1"); NYDFSRefs=@("500.07","500.12"); ImplGuide=MakeUrl $REM_AMS "AMS-01" },
    @{ ControlNumber="AMS-02"; ControlName="User account lifecycle management processes are documented and enforced."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("1"); NYDFSRefs=@("500.07","500.13"); ImplGuide=MakeUrl $REM_AMS "AMS-02" },
    @{ ControlNumber="AMS-03"; ControlName="Audit logging is enabled and retained for all identity and access events."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("1","11"); NYDFSRefs=@("500.06","500.14"); ImplGuide=MakeUrl $REM_AMS "AMS-03" },
    @{ ControlNumber="AMS-04"; ControlName="MFA is enforced for all user accounts within the AMS environment."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("1"); NYDFSRefs=@("500.12"); ImplGuide=MakeUrl $REM_AMS "AMS-04" },
    @{ ControlNumber="AMS-05"; ControlName="Administrative accounts are restricted to privileged tasks only and reviewed periodically."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=2; Policies=@("1"); NYDFSRefs=@("500.07","500.01(n)p"); ImplGuide=MakeUrl $REM_AMS "AMS-05" },
    @{ ControlNumber="AMS-06"; ControlName="Data encryption is enforced at rest and in transit."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("7"); NYDFSRefs=@("500.15a","500.15b"); ImplGuide=MakeUrl $REM_AMS "AMS-06" },

    # EXT Controls
    @{ ControlNumber="EXT-01"; ControlName="A CAA DNS record is published to restrict unauthorized certificate issuance."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("6"); NYDFSRefs=@("500.05","500.08"); ImplGuide=MakeUrl $REM_EXT "EXT-01" },
    @{ ControlNumber="EXT-02"; ControlName="A DMARC policy is published and enforced to prevent email domain spoofing."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("6"); NYDFSRefs=@("500.08","500.15a"); ImplGuide=MakeUrl $REM_EXT "EXT-02" },
    @{ ControlNumber="EXT-03"; ControlName="The robots.txt file is configured to prevent unintended exposure of sensitive paths."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("4"); NYDFSRefs=@("500.08"); ImplGuide=MakeUrl $REM_EXT "EXT-03" },
    @{ ControlNumber="EXT-04"; ControlName="Security headers (CSP, HSTS, X-Frame-Options, etc.) are configured on all web properties."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("4","8"); NYDFSRefs=@("500.08"); ImplGuide=MakeUrl $REM_EXT "EXT-04" },
    @{ ControlNumber="EXT-05"; ControlName="An SPF record is published to authorize legitimate mail senders for the domain."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=1; Policies=@("6"); NYDFSRefs=@("500.08"); ImplGuide=MakeUrl $REM_EXT "EXT-05" },

    # WA Controls (Web Application scanner)
    @{ ControlNumber="WA-1.1"; ControlName="Web application authentication and session controls protect credentials and prevent session hijacking."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=6; Policies=@("1"); NYDFSRefs=@("500.07","500.12"); ImplGuide=MakeUrl $REM_WA "WA-1.1" },
    @{ ControlNumber="WA-1.2"; ControlName="Web application transport security enforces strong TLS, valid certificates, HSTS, and HTTPS-only delivery."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=4; Policies=@("7","8"); NYDFSRefs=@("500.15a","500.15b"); ImplGuide=MakeUrl $REM_WA "WA-1.2" },
    @{ ControlNumber="WA-1.3"; ControlName="Web application security headers harden the browser against XSS, clickjacking, and MIME confusion attacks."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=6; Policies=@("4","8"); NYDFSRefs=@("500.08"); ImplGuide=MakeUrl $REM_WA "WA-1.3" },
    @{ ControlNumber="WA-1.4"; ControlName="Web application configuration does not disclose stack versions, source files, backup files, or directory listings."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=5; Policies=@("8"); NYDFSRefs=@("500.08"); ImplGuide=MakeUrl $REM_WA "WA-1.4" },
    @{ ControlNumber="WA-1.5"; ControlName="Web application input handling resists CSRF, reflected XSS, and SQL injection."; Domain="Technology"; Applicable="Applicable"; EvidenceStatus="Not Started"; InherentRisk=4; Policies=@("8"); NYDFSRefs=@("500.08"); ImplGuide=MakeUrl $REM_WA "WA-1.5" }
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

function Find-CustomerRecId { param([string]$Name)
    $resp = Invoke-SS -Method POST -Url "$BaseUrl/$($AppId.Customers)/records/list/" -Body @{ filter=@{}; sort=@(); fields_to_search=@() }
    $match = $resp.items | Where-Object { $_.title -eq $Name }
    if (-not $match) { throw "Customer '$Name' not found in SmartSuite." }
    return $match.id }

# =============================================================================
# MAIN
# =============================================================================
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
        $F.RA_Title=$("$CustomerName - $($ctrl.ControlNumber)")
        $F.RA_LinkCAM=@($CAMRecId)
        $F.RA_ControlNumber=$ctrl.ControlNumber
        $F.RA_ControlName=$ctrl.ControlName
        $F.RA_Domain=$ctrl.Domain
        $F.RA_Applicable=$ctrl.Applicable
        $F.RA_EvidenceStatus=$ctrl.EvidenceStatus
        $F.RA_InherentRisk=$ctrl.InherentRisk
        $F.RA_AssessYear=$AssessmentYear
    }
    if ($null -ne $ctrl.ImplGuide) { $raBody[$F.RA_ImplGuide] = $ctrl.ImplGuide }
    if ($ctrl.Policies -and $ctrl.Policies.Count -gt 0) {
        $raBody[$F.RA_LinkPolicies] = @(Get-PolicyRecIDs -PolicyNums $ctrl.Policies) }
    if ($ctrl.NYDFSRefs -and $ctrl.NYDFSRefs.Count -gt 0) {
        $raBody[$F.RA_LinkNYCRR] = @(Get-NYCRRRecIDs -Sections $ctrl.NYDFSRefs) }
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
    Write-Host "   OK: Policy $spId updated" }

Write-Host ""; Write-Host ">> Adding customer to all $($TrainingVideoRecIDs.Count) Training and Awareness Video records"
foreach ($tvId in $TrainingVideoRecIDs) {
    Update-Record -AppId $AppId.TrainingVideos -RecId $tvId -Fields @{ $F.TV_Customers=@($CustomerRecId) } | Out-Null
    Write-Host "   OK: Video $tvId updated" }

Write-Host ""; Write-Host ">> Updating Customer record with all links"
Update-Record -AppId $AppId.Customers -RecId $CustomerRecId -Fields @{
    $F.Cust_CAM=@($CAMRecId); $F.Cust_CRB=@($CRBRecId)
    $F.Cust_RecoveryPlans=@($RPRecId); $F.Cust_RiskAssess=$RARecIds.ToArray()
    $F.Cust_SecPolicies=$SecurityPolicyRecIDs; $F.Cust_TrainVideos=$TrainingVideoRecIDs
} | Out-Null
Write-Host "   OK: Customer record updated"

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
```