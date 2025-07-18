function routeAPICallback(reportType, response) {
  Logger.log('Begin display data');

  var sheetName;

  switch (reportType) {    
    case "account_totals":
      sheetName = 'Account Totals';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAccountTotals(response, sheetName);
      break;
    case "additional_fees":
      sheetName = 'Additional Fees';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAdditionalFees(response, sheetName);
      break;
    case "affordable_housing_hud_waitlist":
      sheetName = 'Affordable Housing HUD Waitlist';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAffordableHousingHudWaitlist(response, sheetName);
      break;
    case "affordable_housing_hud_waitlist_activities":
      sheetName = 'Affordable Housing HUD Waitlist Activities';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAffordableHousingHudWaitlistActivities(response, sheetName);
      break;
    case "affordable_housing_program_status":
      sheetName = 'Affordable Housing Program Status';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAffordableHousingProgramStatus(response, sheetName);
      break;
    case "affordable_housing_tenant_demographic":
      sheetName = 'Affordable Housing Tenant Demographic';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAffordableHousingTenantDemographic(response, sheetName);
      break;
    case "affordable_housing_unit_directory":
      sheetName = 'Affordable Housing Unit Directory';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAffordableHousingUnitDirectory(response, sheetName);
      break;
    case "aged_payables_summary":
      sheetName = 'Aged Payables Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAgedPayablesSummary(response, sheetName);
      break;
    case "aged_receivables_detail":
      sheetName = 'Aged Receivables Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAgedReceivableDetail(response, sheetName);
      break;
    case "amenities_by_property":
      sheetName = 'Amenities by Property';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAmenitiesByProperty(response, sheetName);
      break;
    case "annual_budget_comparative":
      sheetName = 'Annual Budget Comparative';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAnnualBudgetComparative(response, sheetName);
      break;
    case "annual_budget_forecast":
      sheetName = 'Annual Budget Forecast';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      hanldeAnnualBudgetForecast(response, sheetName);
      break;
    case "appfolio_stack_usage":
      sheetName = 'AppFolio Stack Usage';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleAppfolioStackUsage(response, sheetName);
      break;
    case "balance_sheet":
      sheetName = 'Balance Sheet';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleBalanceSheet(response, sheetName);
      break;
    case "balance_sheet_comparative":
      sheetName = 'Balance Sheet Comparative';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleBalanceSheetComparative(response, sheetName);
      break;
    case "balance_sheet_comparison":
      sheetName = 'Balance Sheet Property Comparison';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleBalanceSheetPropertyComparison(response, sheetName);
      break;
    case "bank_account_association":
      sheetName = 'Bank Account Association';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleBankAccountAssociation(response, sheetName);
      break;
    case "bill_detail":
      sheetName = 'Bill Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleBillDetail(response, sheetName);
      break;
    case "budget_comparative":
      sheetName = 'Budget Comparative';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleBudgetComparative(response, sheetName);
      break;
    case "budget_comparison":
      sheetName = 'Budget Property Comparison';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleBudgetPropertyComparison(response, sheetName);
      break;
    case "cancelled_processes":
      sheetName = 'Cancelled Workflows';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleCancelledWorkflows(response, sheetName);
      break;
    case "cash_flow":
      sheetName = 'Cash Flow';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleCashFlow(response, sheetName);
      break;
    case "cash_flow_comparison":
      sheetName = 'Cash Flow Property Comparison';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleCashFlowPropertyComparison(response, sheetName);
      break;
    case "cash_flow_detail":
      sheetName = 'Cash Flow Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleCashFlowDetail(response, sheetName);
      break;
    case "charge_detail":
      sheetName = 'Charge Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleChargeDetail(response, sheetName);
      break;
    case "chart_of_accounts":
      sheetName = 'Chart of Accounts';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleChartOfAccounts(response, sheetName);
      break;
    case "check_register":
      sheetName = 'Check Register';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleCheckRegister(response, sheetName);
      break;
    case "check_register_detail":
      sheetName = 'Check Register Detail (Enhanced)';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleCheckRegisterDetailEnhanced(response, sheetName);
      break;
    case "completed_processes":
      sheetName = 'Completed Workflows';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleCompletedWorkflows(response, sheetName);
      break;
    case "delinquency":
      sheetName = 'Delinquency';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleDelinquency(response, sheetName);
      break;
    case "delinquency_as_of":
      sheetName = 'Delinquency As Of';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleDelinquencyAsOf(response, sheetName);
      break;
    case "deposit_register":
      sheetName = 'Deposit Register';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleDepositRegister(response, sheetName);
      break;
    case "eligible_debt_summary":
      sheetName = 'Eligible Debt Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleEligibleDebtSummary(response, sheetName);
      break;
    case "email_delivery_errors":
      sheetName = 'Email Delivery Errors';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleEmailDeliveryErrors(response, sheetName);
      break;
    case "expense_distribution":
      sheetName = 'Expense Distribution';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleExpenseDistribution(response, sheetName);
      break;
    case "expense_register":
      sheetName = 'Expense Register';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleExpenseRegister(response, sheetName);
      break; 
    case "fixed_assets":
      sheetName = 'Fixed Assets';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleFixedAssets(response, sheetName);
      break;
    case "flows_assigned_tasks":
      sheetName = 'Realm-X Flows Assigned Tasks';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleRealmXFlowsAssignedTasks(response, sheetName);
      break;
    case "general_ledger":
      sheetName = 'General Ledger';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleGeneralLedger(response, sheetName);
      break;
    case "gross_potential_rent_enhanced":
      sheetName = 'Gross Potential Rent';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleGrossPotentialRent(response, sheetName);
      break;
    case "guest_card_inquiries":
      sheetName = 'Guest Card Inquiries';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleGuestCardInquiries(response, sheetName);
      break;
    case "guest_cards":
      sheetName = 'Guest Card Interests';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleGuestCardInterests(response, sheetName);
      break;
    case "import_variances":
      sheetName = 'Import Variances';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleImportVariances(response, sheetName);
      break;
    case "in_progress_workflows":
      sheetName = 'In-Progress Workflows';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleInProgressWorkflows(response, sheetName);
      break;
    case "inactive_guest_cards":
      sheetName = 'Inactive Guest Card Interests';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleInactiveGuestCardInterests(response, sheetName);
      break;
    case "income_register":
      sheetName = 'Income Register';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleIncomeRegister(response, sheetName);
      break;
    case "income_statement":
      sheetName = 'Income Statement';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleIncomeStatement(response, sheetName);
      break;
    case "income_statement_comparative":
      sheetName = 'Income Statement Comparative';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleIncomeStatementComparative(response, sheetName);
      break;
    case "income_statement_comparison":
      sheetName = 'Income Statement Property Comparison';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleIncomeStatementPropertyComparison(response, sheetName);
      break;
    case "income_statement_date_range":
      sheetName = 'Income Statement Date Range';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleIncomeStatementDateRange(response, sheetName);
      break;
    case "insurance_usage":
      sheetName = 'Insurance Usage';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleInsuranceUsage(response, sheetName);
      break;
    case "inventory_status":
      sheetName = 'Inventory Status';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleInventoryStatus(response, sheetName);
      break;
    case "inventory_usage":
      sheetName = 'Inventory Usage';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleInventoryUsage(response, sheetName);
      break;
    case "keys_detail":
      sheetName = 'Keys Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleKeysDetail(response, sheetName);
      break;
    case "late_fee_policy_comparison":
      sheetName = 'Late Fee Policy Comparison';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleLateFeePolicyComparison(response, sheetName);
      break;
    case "lease_expiration_detail":
      sheetName = 'Lease Expiration Detail by Month';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleLeaseExpirationDetailByMonth(response, sheetName);
      break;
    case "lease_expiration_summary":
      sheetName = 'Lease Expiration Summary by Month';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleLeaseExpirationSummaryByMonth(response, sheetName);
      break;
    case "lease_history":
      sheetName = 'Lease History';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleLeaseHistory(response, sheetName);
      break;
    case "leasing_agent_performance":
      sheetName = 'Leasing Agent Performance';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleLeasingAgentPerformance(response, sheetName);
      break;
    case "leasing_funnel_performance":
      sheetName = 'Leasing Funnel Performance';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleLeasingFunnelPerformance(response, sheetName);
      break;
    case "leasing_summary":
      sheetName = 'Leasing Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleLeasingSummary(response, sheetName);
      break;
    case "loans":
      sheetName = 'Loans';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleLoans(response, sheetName);
      break;
    case "occupancy_custom_fields":
      sheetName = 'Occupancy Custom Fields';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleOccupancyCustomFields(response, sheetName);
      break;
    case "occupancy_summary":
      sheetName = 'Occupancy Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleOccupancySummary(response, sheetName);
      break;
    case "owner1099":
      sheetName = 'Owner 1099 Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleOwner1099Summary(response, sheetName);
      break;
    case "owner1099_detail":
      sheetName = 'Owner 1099 Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleOwner1099Detail(response, sheetName);
      break;
    case "owner_custom_fields":
      sheetName = 'Owner Custom Fields';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleOwnerCustomFields(response, sheetName);
      break;
    case "owner_directory":
      sheetName = 'Owner Directory';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleOwnerDirectory(response, sheetName);
      break;
    case "owner_leasing":
      sheetName = 'Owner Leasing';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleOwnerLeasing(response, sheetName);
      break;
    case "owner_withholdings":
      sheetName = 'Owner Withholdings';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleOwnerWithholdings(response, sheetName);
      break;
    case "payment_plans":
      sheetName = 'Payment Plans';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handlePaymentPlans(response, sheetName);
      break;
    case "premium_leads_billing_detail":
      sheetName = 'Premium Listing Billing Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handlePremiumListingBillingDetail(response, sheetName);
      break;
    case "project_budget_detail":
      sheetName = 'Project Budget Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleProjectBudgetDetail(response, sheetName);
      break;
    case "property_budget":
      sheetName = 'Property Budget';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handlePropertyBudget(response, sheetName);
      break;
    case "property_custom_fields":
      sheetName = 'Property Custom Fields';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handlePropertyCustomFields(response, sheetName);
      break;
    case "property_directory":
      sheetName = 'Property Directory';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handlePropertyDirectory(response, sheetName);
      break;
    case "property_group_directory":
      sheetName = 'Property Group Directory';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handlePropertyGroupDirectory(response, sheetName);
      break;
    case "property_performance":
      sheetName = 'Property Performance';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handlePropertyPerformance(response, sheetName);
      break;
    case "property_staff_assignments":
      sheetName = 'Property Staff Assignments';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handlePropertyStaffAssignments(response, sheetName);
      break;
    case "prospect_source_tracking":
      sheetName = 'Prospect Source Tracking';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleProspectSourceTracking(response, sheetName);
      break;
    case "purchase_order":
      sheetName = 'Purchase Order';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handlePurchaseOrder(respone, sheetName);
      break;
    case "receivables_activity":
      sheetName = 'Receivables Activity';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleReceivablesActivity(response, sheetName);
      break;
    case "renewal_summary":
      sheetName = 'Renewal Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleRenewalSummary(response, sheetName);
      break;
    case "rent_roll":
      sheetName = 'Rent Roll';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleRentRoll(response, sheetName);
      break;
    case "rent_roll_commercial":
      sheetName = 'Rent Roll Commercial';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleRentRollCommercial(response, sheetName);
      break;
    case "rent_roll_itemized":
      sheetName = 'Rent Roll Itemized';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleRentRollItemized(response, sheetName);
      break;
    case "rentable_items":
      sheetName = 'Rentable Items';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleRentableItems(response, sheetName);
      break;
    case "rental_applications":
      sheetName = 'Rental Applications';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleRentalApplications(response, sheetName);
      break;
    case "resident_financial_activity":
      sheetName = 'Resident Financial Activity';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleResidentFinancialActivity(response, sheetName);
      break;
    case "screening_assessments":
      sheetName = 'Screening Assessments';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleScreeningAssessments(response, sheetName);
      break;
    case "screening_usage":
      sheetName = 'Screening Usage';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleScreeningUsage(response, sheetName);
      break;
    case "security_deposit_funds_detail":
      sheetName = 'Security Deposit Funds Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleSecurityDepositFundsDetail(response, sheetName);
      break;
    case "showings":
      sheetName = 'Showings';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleShowings(response, sheetName);
      break;
    case "surveys_summary":
      sheetName = 'Survey Responses';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleSurveyResponses(response, sheetName);
      break;
    case "tax_credit_building_directory":
      sheetName = 'Tax Credit Building Directory';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTaxCreditBuildingDirectory(response, sheetName);
      break;
    case "tenant_debt_collections_status":
      sheetName = 'Debt Collections Status';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleDebtCollectionsStatus(response, sheetName);
      break;
    case "tenant_directory":
      sheetName = 'Tenant Directory';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTenantDirectory(response, sheetName);
      break;
    case "tenant_ledger":
      sheetName = 'Tenant Ledger';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTenantLedger(response, sheetName);
      break;
    case "tenant_tickler":
      sheetName = 'Tenant Tickler';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTenantTickler(response, sheetName);
      break;
    case "tenant_transactions_summary":
      sheetName = 'Tenant Transactions Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTenantTransactionsSummary(response, sheetName);
      break;
    case "tenant_unpaid_charges_summary":
      sheetName = 'Tenant Unpaid Charges Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTenantUnpaidChargesSummary(response, sheetName);
      break;
    case "tenant_vehicle_info":
      sheetName = 'Tenant Vehicle Info';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTenantVehicleInfo(response, sheetName);
      break;
    case "trial_balance":
      sheetName = 'Trial Balance';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTrialBalance(response, sheetName);
      break;
    case "trial_balance_by_property":
      sheetName = 'Trial Balance by Property';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTrialBalanceByProperty(response, sheetName);
      break;
    case "trust_account_balance":
      sheetName = 'Trial Account Balance';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTrailAccountBalance(response, sheetName);
      break;
    case "trust_account_balance_detail":
      sheetName = 'Trial Account Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleTrustAccountDetail(response, sheetName);
      break;
    case "twelve_month_cash_flow":
      sheetName = 'Cash Flow 12 Month';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleCashFlow12Month(response, sheetName);
      break;
    case "twelve_month_income_statement":
      sheetName = 'Income Statement 12 Month';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleIncomeStatement12Month(response, sheetName);
      break;
    case "unit_custom_fields":
      sheetName = 'Unit Custom Fields';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleUnitCustomFields(response, sheetName);
      break;
    case "unit_directory":
      sheetName = 'Unit Directory';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleUnitDirectory(response, sheetName);
      break;
    case "unit_inspection":
      sheetName = 'Unit Inspection';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleUnitInspection(response, sheetName);
      break;
    case "unit_turn_detail":
      sheetName = 'Unit Turn Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleUnitTurnDetail(response, sheetName);
      break;
    case "unit_vacancy":
      sheetName = 'Unit Vacancy Detail';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleUnitVacancyDetail(response, sheetName);
      break;
    case "unpaid_balances_by_month":
      sheetName = 'Unpaid Balances by Month';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleUnpaidBalancesByMonth(response, sheetName);
      break;
    case "upcoming_activities":
      sheetName = 'Activities Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleActivitiesSummary(response, sheetName);
      break;
    case "vendor1099":
      sheetName = 'Vendor 1099 Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleVendor1099Summary(response, sheetName);
      break;
    case "vendor_custom_fields":
      sheetName = 'Vendor Custom Fields';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleVendorCustomFields(response, sheetName);
      break;
    case "vendor_directory":
      sheetName = 'Vendor Directory';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleVendorDirectory(response, sheetName);
      break;
    case "vendor_ledger":
      sheetName = 'Vendor Ledger';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleVendorLedger(response, sheetName);
      break;
    case "work_order":
      sheetName = 'Work Order';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleWorkOrder(response, sheetName);
      break;
    case "work_order_labor_summary":
      sheetName = 'Work Order Labor Summary';
      sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
      if (!sheet) {
        sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      }
      handleWorkOrderLaborSummary(response, sheetName);
      break;
  }
}
