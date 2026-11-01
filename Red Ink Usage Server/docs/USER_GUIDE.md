# Red Ink Usage Server — User Guide

**Audience:** Authorized staff who review usage figures and reports. Configuration, security, installation and token management are described in the [Administrator Guide](ADMINISTRATOR_GUIDE.md).

## Access and sign-in

Open the URL supplied by your administrator (for example `http://localhost:8765` when the server runs on the same PC, or an organization-approved HTTPS address). Sign in with your administrator-issued credentials. Some local installations offer a passwordless sign-in; it is unavailable for remote administration. Keep token-bearing configuration links private.

## Monthly and daily usage

**Monthly users** counts distinct observed license identities for a product and license scope in a given month. One user opening Word, Excel and Outlook counts **once**, not three times. **Daily users** counts distinct identities with an explicitly recorded add-in load on that client-local day. Daily totals must not be added to calculate billable monthly users. Late offline records belong to the original observed dates, not the date of arrival.

Use the date/product/scope filters to select the records you need. A report is only as complete as the evidence received by the server; devices that have not reported yet will not appear. Historical report snapshots reflect the data available when they were created.

## Feature and agent activity

The feature dashboard uses **LogPath** event records, which are separate from license evidence. The activity chart shows direct user commands and explicitly recorded top-level agent-run starts; child tool calls are not additional runs. You can review common features and the frequency of actual tools, skills and delegated agents separately. Hover over or focus a feature group to inspect the commands within it.

An activity count indicates an observed action, **not** its successful completion, duration, time saved or productivity. Older add-ins may lack run-start markers, so their agent-run totals may be unavailable. Word, Excel and Outlook are labelled separately where the evidence identifies the host.

## User statistics and privacy

Where the source log includes a stable pseudonymous identifier, the dashboard shows activity by that identifier. Figures include activity per calendar day, per seven calendar days, per active day and per active week. The denominators differ: calendar rates include inactive days; active-day/week averages do not. Peaks and medians help distinguish occasional intense use from regular use.

`LicenseCounterAnon=False` can allow recoverable license user identifiers in **license reports**, but does not automatically link them to LogPath user IDs. The server does not invent this linkage. Follow your organization's privacy and access-control rules when exporting personal usage details.

## Reports and exports

Use the report filters before requesting a printable report, CSV or JSON export. Review the product, scope, reporting period and snapshot date before using the figures for invoices or internal reviews. Printable reports are designed for a light background. Protect exported files according to your organization’s retention policy.

## When something looks wrong

If a user is missing, first check the reporting period and whether their add-in actually sent data. If tool counts seem higher than runs, remember that one run may use several tools; do not add tool counts into run totals. If a dashboard says a log source is pending, later scans may complete its import. Report persistent warnings or access errors to your server administrator rather than modifying the source files.

The administrator maintains the server, configures publishing/collection URLs, rotates tokens, controls retention and starts or stops the service. Ordinary report viewers should not make those changes.
