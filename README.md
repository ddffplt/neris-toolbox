# neris-toolbox
Scripts that help you access and use your NERIS data. 

These scripts leverage the NERIS API using the NERIS API Client (https://github.com/ulfsri/neris-api-client). 

## SFMO query scopes

SFMO query reports offer the entire state, a Massachusetts county, one entity,
or multiple entities. County selection uses the department's FD25 county ID
prefix (for example `FD25017` or `FD25017*`), not the incident location.
Enter multiple entity IDs or names separated by commas. Some incident scripts
default to Massachusetts unless a state is supplied with `--state`; their
existing command-line options remain available.

The aid activity report temporarily skips the separate statewide collection for
the **Aid Given In Lieu Of** incident-detail tab; that tab displays a notice.
The **Summary** and **Entity Aid Reports** tabs still include all aid category
counts, including aid given in lieu of, from the selected entities' incidents.
When the entity list returns HTTP 503 or 504, the report retries the affected
page up to five times with increasing delays. If retrieval still fails, the
report stops without creating a workbook; try again when the service recovers.
