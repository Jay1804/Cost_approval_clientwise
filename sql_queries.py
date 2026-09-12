"""
SQL query templates used by the Cost Approval Tracker.

These mirror exactly the queries that used to be run manually through the
Query Browser / MIS export portal ("CS Hub" -> ... and "Extra Query - Bridge"
-> Client wise Flexi field names query). Client filtering is applied by
appending an extra `AND ecm.client_id IN (...)` / `AND eccf.client_id IN (...)`
clause so a single client (or a handful of clients) can be pulled without
scanning the whole table.
"""

# Clients excluded everywhere (internal/test/dummy accounts)
EXCLUDED_CLIENT_IDS = (1000, 88031, 151588, 263464, 263586, 231747)

_EXCLUDED_CLIENT_IDS_SQL = ",".join(str(c) for c in EXCLUDED_CLIENT_IDS)


def _client_filter_sql(client_ids, column):
    """Build an `AND column IN (...)` clause for a validated list of int client ids."""
    if not client_ids:
        return ""
    ids = [int(c) for c in client_ids]
    return f" AND {column} IN ({','.join(str(i) for i in ids)})"


# ---------------------------------------------------------------------------
# Lightweight client list, used to populate the client-selection dropdown.
# Same filters as the main query, but only returns distinct client id/name.
# ---------------------------------------------------------------------------
CLIENT_LIST_QUERY = f"""
SELECT DISTINCT ecm.client_id, emc.Company_name
FROM ec_case_master ecm
LEFT JOIN ec_case_checks ecc ON ecm.case_id = ecc.case_id
LEFT JOIN ec_master_company emc ON emc.company_id = ecm.client_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id = ecm.candidate_id
WHERE ecc.check_status = 3
  AND ecm.case_status NOT IN (6,8,14)
  AND ecm.client_id NOT IN ({_EXCLUDED_CLIENT_IDS_SQL})
  AND ecc1.first_name NOT LIKE '%dummy%' AND ecc1.first_name NOT LIKE '%test%'
  AND (ecc.insuff_remarks LIKE '%cost%' OR ecc.insuff_remarks LIKE '%approval%')
ORDER BY emc.Company_name
"""


def get_case_data_query(client_ids=None):
    """Main 'Client wise all cases with all checks with unique check name-Insuff - Cost Approval' query."""
    return f"""
SELECT ecc.Case_Check_id,
	ecm.client_id,
	case_ars_no,
	emc.Company_name,
	CONCAT(first_name,' ',Middle_name,' ',last_name) AS Candidate_name,
	Process_name,
	office_name location,
	received_date AS case_received_date,

	checkpoint_live.fn_case_status(case_status) AS 'Case_status',
	ecc.check_name AS 'check_name',

	(SELECT insuff_raised_date FROM ec_case_insuff_details ecid
	     WHERE ecc.case_check_id=ecid.case_check_id AND insuff_raised_date IS NOT NULL ORDER BY insuff_id DESC LIMIT 1) AS 'Insuff Raised Date',
	(SELECT insufficiency FROM ec_case_insuff_details ecid
	     WHERE ecc.case_check_id=ecid.case_check_id AND insuff_raised_date IS NOT NULL ORDER BY insuff_id DESC LIMIT 1) AS 'insuff_remarks',
	(SELECT insuff_fulfil_date FROM ec_case_insuff_details ecid
	     WHERE ecc.case_check_id=ecid.case_check_id ORDER BY insuff_id DESC LIMIT 1) AS 'insuff_fulfill_date',
	(SELECT cost FROM ec_case_insuff_details ecid
	     WHERE ecc.case_check_id=ecid.case_check_id AND insuff_raised_date IS NOT NULL ORDER BY insuff_id DESC LIMIT 1) AS 'Cost',
	checkpoint_live.fn_check_status(check_status) AS 'check_status',
	check_severity,
	closure_comments,
	ec.check_name AS 'Check_unique_name',
	fn_vs_name(ecc.case_check_id) AS 'VS Name',

IF(ecc.family_id=5,emc4.city_name,
	  IF(ecc.family_id=4,emc6.city_name,'')) AS 'VS City',

IF(ecc.family_id=5,ems.state_name,
	  IF(ecc.family_id=4,ems1.state_name,'')) AS 'VS State',

IF(ecc.family_id=5,emc5.country_name,
	  IF(ecc.family_id=4,emc7.country_name,'')) AS 'VS Country',

ecf.case_id,ecf.CASE_FLEX_FIELD1,ecf.CASE_FLEX_FIELD2,ecf.CASE_FLEX_FIELD3,ecf.CASE_FLEX_FIELD4,
ecf.CASE_FLEX_FIELD5,ecf.CASE_FLEX_FIELD6,ecf.CASE_FLEX_FIELD7,ecf.CASE_FLEX_FIELD8,
ecf.CASE_FLEX_FIELD9,ecf.CASE_FLEX_FIELD10,ecf.CASE_FLEX_FIELD11,ecf.CASE_FLEX_FIELD12,
ecf.CASE_FLEX_FIELD13,ecf.CASE_FLEX_FIELD14,ecf.CASE_FLEX_FIELD15,ecf.CASE_FLEX_FIELD16,
ecf.CASE_FLEX_FIELD17,ecf.CASE_FLEX_FIELD18,ecf.CASE_FLEX_FIELD19,ecf.CASE_FLEX_FIELD20,
ecf.CASE_FLEX_FIELD21,ecf.CASE_FLEX_FIELD22,ecf.CASE_FLEX_FIELD23,ecf.CASE_FLEX_FIELD24,
ecf.CASE_FLEX_FIELD25,ecf.CASE_FLEX_FIELD26,ecf.CASE_FLEX_FIELD27,ecf.CASE_FLEX_FIELD28,
ecf.CASE_FLEX_FIELD29,ecf.CASE_FLEX_FIELD30

FROM ec_case_master ecm
LEFT JOIN ec_case_checks ecc ON ecm.case_id = ecc.case_id
LEFT JOIN ec_master_company emc ON emc.company_id = ecm.client_id
LEFT JOIN ec_client_process ecp ON ecm.process_id=ecp.process_id
LEFT JOIN ec_case_candidates ecc1 ON ecc1.candidate_id=ecm.candidate_id
LEFT JOIN ec_master_company_locations emcl ON ecm.client_office_id=emcl.office_id
LEFT JOIN ec_user_details eud ON ecc.check_verifier=eud.user_id

LEFT JOIN ec_case_check_verification_source eccvs ON ecc.case_check_id=eccvs.case_check_id
LEFT JOIN ec_master_educational_institute emei ON eccvs.org_id=emei.institute_id
LEFT JOIN ec_master_company emc1 ON eccvs.org_id=emc1.company_id
LEFT JOIN ec_master_city emc2 ON eccvs.org_id=emc2.city_id
LEFT JOIN ec_master_country emc3 ON emc2.country_id=emc3.country_id
LEFT JOIN ec_master_city emc4 ON emc1.city_id=emc4.city_id
LEFT JOIN ec_master_state ems ON emc1.state_id=ems.state_id
LEFT JOIN ec_master_country emc5 ON emc1.country_id=emc5.country_id

LEFT JOIN ec_master_city emc6 ON emei.city_id=emc6.city_id
LEFT JOIN ec_master_state ems1 ON emei.state_id=ems1.state_id
LEFT JOIN ec_master_country emc7 ON emei.country_id=emc7.country_id
LEFT JOIN ec_case_fields ecf on ecm.case_id=ecf.case_id

LEFT JOIN ec_checks ec ON ecc.check_id=ec.check_id
WHERE ecc.check_status=3 AND case_status NOT IN (6,8,14) AND ecm.client_id NOT IN ({_EXCLUDED_CLIENT_IDS_SQL})
AND first_name NOT LIKE '%dummy%' AND first_name NOT LIKE '%test%'
AND (ecc.insuff_remarks LIKE '%cost%' OR ecc.insuff_remarks LIKE '%approval%')
{_client_filter_sql(client_ids, 'ecm.client_id')}
"""


def get_flexi_field_query(client_ids=None):
    """'Client wise Flexi field names query' - maps CASE_FLEX_FIELDx to their client-specific names."""
    return f"""
SELECT eccf.client_id,
       client_external_id,
       company_name AS 'Client Name',
       field_id,
       field_name,
       IF(`status`=0,'Passive','Active') AS 'Status'
FROM ec_client_case_fields eccf
LEFT JOIN ec_client ec ON eccf.client_id=ec.client_id
INNER JOIN ec_master_company emc ON eccf.client_id=emc.company_id
WHERE 1=1
{_client_filter_sql(client_ids, 'eccf.client_id')}
"""
