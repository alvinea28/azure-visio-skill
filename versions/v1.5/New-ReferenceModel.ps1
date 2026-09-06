[CmdletBinding()]
param(
    [Parameter(Mandatory)][string]$OutputPath,
    [ValidateRange(2,3)][int]$SpokeCount = 2
)
Set-StrictMode -Version 2
$ErrorActionPreference = 'Stop'
if ($OutputPath -notmatch '^(?:[A-Za-z]:\\|\\\\[^\\]+\\[^\\]+(?:\\|$))' -or [IO.Path]::GetExtension($OutputPath) -ine '.json') {
    throw 'OutputPath must be an absolute .json path.'
}
if (Test-Path -LiteralPath $OutputPath) { throw "Refusing to overwrite model: $OutputPath" }
$pages = [System.Collections.Generic.List[object]]::new()
$script:nodes = $null; $script:edges = $null
$sources = @{
    landing = 'https://learn.microsoft.com/en-us/azure/cloud-adoption-framework/ready/landing-zone/'
    network = 'https://learn.microsoft.com/en-us/azure/architecture/networking/architecture/hub-spoke'
    dns = 'https://learn.microsoft.com/en-us/azure/private-link/private-endpoint-dns-integration'
    firewall = 'https://learn.microsoft.com/en-us/azure/firewall/overview'
    pe = 'https://learn.microsoft.com/en-us/azure/private-link/private-endpoint-overview'
    governance = 'https://learn.microsoft.com/en-us/azure/cloud-adoption-framework/ready/landing-zone/design-area/resource-org'
    monitor = 'https://learn.microsoft.com/en-us/azure/azure-monitor/fundamentals/overview'
}
function Start-Page([string]$Name, [string]$Title, [string]$Subtitle) {
    $script:nodes = [System.Collections.Generic.List[object]]::new()
    $script:edges = [System.Collections.Generic.List[object]]::new()
    $pages.Add([ordered]@{ name=$Name; title=$Title; subtitle=$Subtitle; width=20; height=11.25; nodes=$script:nodes; edges=$script:edges })
}
function Node([string]$Id, [string]$Kind, [string]$Text, [double]$X, [double]$Y, [double]$W, [double]$H,
              [string]$Icon='', [string]$Parent='', [string]$Component='', [string]$Url='',
              [int]$Size=12, [string]$Color='RGB(100,116,139)', [string]$Layer='Architecture') {
    if (-not $Component) { $Component = $Id }
    $script:nodes.Add([ordered]@{
        id=$Id; kind=$Kind; label=$Text; x=$X; y=$Y; width=$W; height=$H
        icon=$Icon; parent=$Parent; component=$Component; url=$Url
        fontSize=$Size; color=$Color; layer=$Layer; state='Illustrative'
    })
}
function Edge([string]$Id, [string]$Source, [string]$Target, [string]$Kind, [string]$Text='',
              [string]$From='', [string]$To='') {
    $e = [ordered]@{id=$Id; source=$Source; target=$Target; kind=$Kind; label=$Text}
    if ($From) { $e.sourceSide=$From }
    if ($To) { $e.targetSide=$To }
    $script:edges.Add($e)
}

Start-Page '01 Overview' 'Azure landing zone | Hub-and-spoke' 'Generic reference environment. One illustrative region. No real tenant, addresses, sizing, or deployment is implied.'
Node 'entra' 'card' "Microsoft Entra ID`nTenant-wide identity, outside subscriptions" 10 9.1 18.6 0.85 'entra' '' 'tenant-identity' $sources.landing
Node 'platform' 'container' 'Platform landing zone | Shared ownership' 3.65 5.5 5.85 5.7 '' '' 'platform' $sources.landing
Node 'applications' 'container' 'Application landing zones | Workload ownership' 13.1 5.5 12.4 5.7 '' '' 'applications' $sources.landing
Node 'connectivity-sub' 'card' "Connectivity subscription`nHub VNet + Azure Firewall" 3.65 6.85 4.8 1.25 'subscription' 'platform' 'connectivity-sub' $sources.network
Node 'dns-zone' 'card' "Azure Private DNS zone`nShared name-resolution configuration" 3.65 5.15 4.8 1.25 'dns' 'platform' 'dns-zone' $sources.dns
Node 'management-sub' 'card' "Management subscription`nCentral logging and monitoring" 3.65 3.45 4.8 1.25 'monitor' 'platform' 'management-sub' $sources.monitor
for ($i=1; $i -le $SpokeCount; $i++) {
    $key = 'spoke-{0:00}' -f $i; $x = 9.1+($i-1)*4
    Node "$key-sub" 'card' "Workload $i subscription`nTeam-owned resources" $x 6.65 3.4 1.15 'subscription' 'applications' "$key-sub" $sources.landing
    Node "$key-vnet" 'card' "Spoke $i virtual network`nIsolated address space" $x 4.9 3.4 1.15 'vnet' 'applications' "$key-vnet" $sources.network
    Edge "overview-$key" "$key-sub" "$key-vnet" 'governance' 'Resource ownership'
}
Node 'overview-guidance' 'note' "Scope: shared connectivity and governance, not a complete enterprise landing zone deployment.`nWorkload subscriptions shown separately. The third column is reserved for additive expansion.`nIdentity infrastructure, hybrid gateways, public ingress, and multi-region recovery are not assumed." 10 1.6 18.6 1.35 '' '' '' $sources.landing 13

Start-Page '02 Governance' 'Governance | Scopes and ownership' 'Illustrative management-group hierarchy. Governance relationships are not network connections.'
Node 'tenant-root' 'card' "Tenant root management group`n<tenant-root>" 10 9.1 5.2 0.95 'managementGroup' '' '' $sources.governance
Node 'platform-mg' 'card' "Platform management group`nPlatform team" 4.1 7.1 4.6 1.05 'managementGroup' '' '' $sources.governance
Node 'workloads-mg' 'card' "Workloads management group`nIllustrative grouping" 13.1 7.1 5.4 1.05 'managementGroup' '' '' $sources.governance
Edge 'gov-root-platform' 'tenant-root' 'platform-mg' 'governance' 'Hierarchy' 'bottom' 'top'
Edge 'gov-root-workloads' 'tenant-root' 'workloads-mg' 'governance' 'Hierarchy' 'bottom' 'top'
Node 'connectivity-sub' 'card' "Connectivity subscription`nHub, firewall, private DNS" 2.15 4.9 3.05 1.2 'subscription' '' 'connectivity-sub' $sources.landing 11
Node 'management-sub' 'card' "Management subscription`nLogging and operations" 5.65 4.9 3.05 1.2 'subscription' '' 'management-sub' $sources.landing 11
Edge 'gov-connectivity' 'platform-mg' 'connectivity-sub' 'governance' '' 'bottom' 'top'
Edge 'gov-management' 'platform-mg' 'management-sub' 'governance' '' 'bottom' 'top'
for ($i=1; $i -le $SpokeCount; $i++) {
    $key='spoke-{0:00}' -f $i; $x=9.1+($i-1)*4
    Node "$key-sub" 'card' "Workload $i subscription`n<subscription-id>" $x 4.9 3.4 1.2 'subscription' '' "$key-sub" $sources.landing
    Edge "gov-$key" 'workloads-mg' "$key-sub" 'governance' '' 'bottom' 'top'
}
Node 'governance-decisions' 'note' "Policy and access decisions`nAssign guardrails at the appropriate management-group or subscription scope.`nUse least privilege and separate platform and workload responsibilities.`nDefine exemptions, deployment identities, budget ownership, and diagnostic requirements." 5.05 2.4 9 2.45 '' '' '' $sources.governance 13
Node 'governance-limits' 'note' "Important boundaries`nMicrosoft Entra ID is tenant-wide, not deployed into an identity subscription.`nAn identity subscription is only needed for actual hosted identity infrastructure.`nThis simplified hierarchy is illustrative, not the full landing zone accelerator hierarchy." 14.95 2.4 9 2.45 '' '' '' $sources.landing 13

Start-Page '03 Network topology' 'Network topology | Regional hub and workload spokes' 'Region: <azure-region> | Address spaces: <non-overlapping-CIDRs> | Peering is nontransitive. Blue links show topology, not inspected traffic.'
Node 'connectivity-sub' 'container' 'Connectivity subscription' 3.55 6.05 5.6 6.65 '' '' 'connectivity-sub' $sources.network
Node 'hub-boundary' 'container' 'Hub virtual network | <hub-CIDR>' 3.55 5.88 5.0 5.5 '' 'connectivity-sub' 'hub-boundary' $sources.network 12 'RGB(37,99,235)'
Node 'hub-vnet' 'card' "Hub VNet`nPeering attachment" 3.55 7.28 4.1 0.9 'vnet' 'hub-boundary' 'hub-vnet' $sources.network
Node 'firewall-subnet' 'container' 'AzureFirewallSubnet | <CIDR, /26 or larger>' 3.55 5.5 4.4 2.25 '' 'hub-boundary' 'firewall-subnet' $sources.firewall 11 'RGB(5,130,94)'
Node 'hub-firewall' 'card' "Azure Firewall`nPrivate IP: <firewall-IP>" 3.55 5.24 3.6 1.0 'firewall' 'firewall-subnet' 'hub-firewall' $sources.firewall
Node 'hub-note' 'note' "Central egress and selected transit.`nNo gateway or hybrid network is assumed." 3.55 3.74 4.3 0.85 '' 'hub-boundary' '' $sources.network 11
for ($i=1; $i -le $SpokeCount; $i++) {
    $key='spoke-{0:00}' -f $i; $x=9.1+($i-1)*4
    Node "$key-sub" 'container' "Workload $i subscription" $x 6.05 3.65 6.65 '' '' "$key-sub" $sources.landing
    Node "$key-boundary" 'container' "Spoke $i VNet | <CIDR>" $x 5.88 3.3 5.5 '' "$key-sub" "$key-boundary" $sources.network 11 'RGB(37,99,235)'
    $vnetY = 7.28
    if ($i -eq 3) { $vnetY = 7.68 }
    Node "$key-vnet" 'card' "Spoke $i VNet`nPeering attachment" $x $vnetY 2.94 0.9 'vnet' "$key-boundary" "$key-vnet" $sources.network 11
    Node "$key-subnet" 'container' 'Workload subnet | <CIDR>' $x 5.8 2.98 1.75 '' "$key-boundary" "$key-subnet" $sources.network 11 'RGB(5,130,94)'
    Node "$key-workload" 'card' "Workload $i`nCompute unspecified" $x 5.52 2.67 0.8 '' "$key-subnet" "$key-workload" $sources.network 11
    if ($i -eq 1) {
        Edge "peer-$key" 'hub-vnet' "$key-vnet" 'peering' "Hub <> Spoke $i" 'right' 'left'
    } else {
        Edge "peer-$key" 'hub-vnet' "$key-vnet" 'peering' "Hub <> Spoke $i" 'top' 'top'
    }
    if ($i -eq 1) {
        Node 'private-subnet' 'container' 'Private endpoint subnet | <CIDR>' $x 3.98 2.98 1.5 '' "$key-boundary" 'private-subnet' $sources.pe 10 'RGB(5,130,94)'
        Node 'blob-endpoint' 'card' "Private endpoint`nStorage blob" $x 3.7 2.67 0.8 'privateLink' 'private-subnet' 'blob-endpoint' $sources.pe 11
    } else {
        Node "$key-route-note" 'note' "Outbound route:`n0.0.0.0/0 -> <firewall-IP>`nNetwork security rules required." $x 3.95 2.9 1.1 '' "$key-boundary" '' $sources.network 11
    }
}
Node 'blob-storage' 'card' "Azure Storage account | blob`nManaged service outside VNets" 9.1 1.66 5.5 1.05 'storage' '' 'blob-storage' $sources.pe
Node 'dns-zone' 'card' "Private DNS zone | centrally owned`nprivatelink.blob.core.windows.net" 16.15 1.66 6 1.05 'dns' '' 'dns-zone' $sources.dns
Edge 'endpoint-service' 'blob-endpoint' 'blob-storage' 'traffic' 'Private Link association' 'bottom' 'top'
Node 'topology-note' 'note' "Private endpoints are in a subnet.`nThe backing Azure service is not.`nSee page 04 for traffic and DNS." 3.45 1.66 5.2 1.05 '' '' '' $sources.pe 12

Start-Page '04 Traffic and DNS' 'Traffic and DNS | Explicit paths, separate relationships' 'Orange arrows carry application traffic. Purple lines describe name resolution and DNS record relationships, not application traffic.'
Node 'egress-band' 'container' 'A. Outbound internet access | Explicit inspected path' 10 8.0 18.6 2.7 '' '' '' $sources.network
Node 'egress-workload' 'card' "Workload 1 subnet`nOutbound traffic" 3 7.73 3.9 1.15 '' 'egress-band' 'spoke-01-workload' $sources.network
Node 'hub-firewall' 'card' "Azure Firewall`nRules + outbound SNAT" 10 7.73 3.9 1.15 'firewall' 'egress-band' 'hub-firewall' $sources.firewall
Node 'internet' 'card' "Allowed public destination`nNo inbound service implied" 17 7.73 3.9 1.15 'internet' 'egress-band' 'internet' $sources.firewall
Edge 'flow-egress-to-firewall' 'egress-workload' 'hub-firewall' 'traffic' "User-defined route`n0.0.0.0/0 -> firewall"
Edge 'flow-firewall-to-internet' 'hub-firewall' 'internet' 'traffic' "Allowed by policy`nSNAT for internet egress"
Node 'private-band' 'container' 'B. Private blob access | Direct private endpoint path in Spoke 1, not firewall inspected in this example' 10 3.56 18.6 5.3 '' '' '' $sources.dns
Node 'private-client' 'card' "Workload 1 client`nUses normal blob hostname" 3 4.72 3.9 1.2 '' 'private-band' 'spoke-01-workload' $sources.dns
Node 'blob-endpoint' 'card' "Private endpoint`nPrivate IP: <endpoint-IP>" 10 4.72 3.9 1.2 'privateLink' 'private-band' 'blob-endpoint' $sources.pe
Node 'blob-storage' 'card' "Azure Storage | blob`nOutside the virtual network" 17 4.72 3.9 1.2 'storage' 'private-band' 'blob-storage' $sources.pe
Edge 'flow-client-endpoint' 'private-client' 'blob-endpoint' 'traffic' 'HTTPS to private IP'
Edge 'flow-endpoint-storage' 'blob-endpoint' 'blob-storage' 'traffic' 'Private Link'
Node 'dns-zone' 'card' "Private DNS zone`nprivatelink.blob.core.windows.net" 10 2.18 5.1 1.15 'dns' 'private-band' 'dns-zone' $sources.dns
Edge 'dns-client-zone' 'private-client' 'dns-zone' 'dns' "Resolution via Azure-provided DNS`nZone linked to requesting VNet" 'bottom' 'left'
Edge 'dns-record-endpoint' 'dns-zone' 'blob-endpoint' 'dns' "Zone group maintains`nA record -> private IP" 'top' 'bottom'
Node 'private-dns-note' 'note' "Central zone ownership does not require`na hub-hosted DNS server.`nLink the zone to each VNet that needs it.`nEndpoint creation alone does not disable`nthe storage public endpoint." 17 2.25 4.5 1.65 '' 'private-band' '' $sources.dns 11

Start-Page '05 Operations and decisions' 'Operations | Ownership, assumptions, and design gaps' 'Design review prompts, not proof of security, availability, or compliance. Service configuration and requirements must be confirmed.'
Node 'assumptions' 'note' "REFERENCE ASSUMPTIONS`n`nOne region; no disaster recovery design.`nSeparate workload subscriptions.`nAzure-provided DNS for clients.`nPrivate zone centrally owned.`nPrivate endpoint shown in Spoke 1.`nNo implicit spoke-to-spoke transit." 3.55 7.75 5.8 3.1 '' '' '' $sources.network 13
Node 'routing-decisions' 'note' "ROUTING AND SECURITY DECISIONS`n`nConfirm CIDRs and non-overlap.`nEnable required peering traffic settings.`nDefine firewall rules and subnet routes.`nUse specific prefixes for inspected transit.`nCheck symmetric return paths.`nPrivate endpoint inspection needs its own design." 10 7.75 5.8 3.1 '' '' '' $sources.network 13
Node 'ownership' 'note' "OWNERSHIP AND OPEN QUESTIONS`n`nPlatform team: hub, firewall, shared DNS.`nWorkload teams: apps, subnets, endpoints.`nDefine public access and admin access.`nChoose sizes, zones, region and budgets.`nDefine recovery objectives and retention.`nHybrid access is a separate extension." 16.45 7.75 5.8 3.1 '' '' '' $sources.landing 13
Node 'hub-firewall' 'card' "Azure Firewall`nDiagnostics enabled by configuration" 3.55 4.87 4.7 1.15 'firewall' '' 'hub-firewall' $sources.firewall
Node 'central-logs' 'card' "Log Analytics workspace`nManagement subscription" 10 4.87 4.7 1.15 'logs' '' 'central-logs' $sources.monitor
Node 'monitor' 'card' "Azure Monitor`nAlerts, dashboards, operations" 16.45 4.87 4.7 1.15 'monitor' '' 'monitor' $sources.monitor
Edge 'telemetry-firewall' 'hub-firewall' 'central-logs' 'telemetry' 'Diagnostic settings'
Edge 'telemetry-monitor' 'central-logs' 'monitor' 'telemetry' 'Queries and alerts'
Node 'ops-guidance' 'note' "VALIDATION BEFORE DEPLOYMENT`nConfirm effective routes, firewall allow/deny behavior, private DNS answers, public access settings, and least-privilege permissions.`nA default route alone does not force more-specific private endpoint or peering traffic through the firewall.`nIf cross-spoke access is required, design explicit routing and inspection. Do not treat peering as transitive." 10 2.99 18.6 1.8 '' '' '' $sources.network 13
Node 'source-landing' 'card' 'Source: Azure landing zones' 3.55 1.35 5.8 0.65 '' '' '' $sources.landing 11
Node 'source-network' 'card' 'Source: Hub-spoke network topology' 10 1.35 5.8 0.65 '' '' '' $sources.network 11
Node 'source-dns' 'card' 'Source: Private endpoint DNS integration' 16.45 1.35 5.8 0.65 '' '' '' $sources.dns 11

$model = [ordered]@{
    schemaVersion=1
    title='Generic Azure hub-and-spoke landing zone'
    description='Illustrative reference; no customer data or Azure resource changes.'
    spokeCount=$SpokeCount
    iconSource='Microsoft Azure stencils installed with desktop Visio; labels updated to current service names.'
    sources=$sources
    pages=$pages
}
[void][IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($OutputPath))
$model | ConvertTo-Json -Depth 20 | Set-Content -LiteralPath $OutputPath -Encoding UTF8
Write-Output $OutputPath
