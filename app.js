require('dotenv').load();
var express = require('express');
var app = express();
app.locals._ = require('lodash');
var _ = require('lodash');
app.locals.moment = require('moment');
var moment = require('moment');

// generate unique filename
var unix_time = moment().unix().toString();
var filename = './parsed_files/' + unix_time + '-' + process.env.FILE_TO_PARSE + '.csv';

var fields_to_use_for_csv = process.env.FIELDS_TO_USE_FOR_CSV;

// csv library
var json2csv = require('json2csv');
var fs = require('fs');
var fields;

if(fields_to_use_for_csv == 1) {
  // Month 1 Fields
  fields = ['response_number', 'submit_date', 'dt_id', 'lead_installer', 'phone', 'dealership', 'respondent_name', 'good_communication', 'good_communication_score', 'satisfaction_with_install', 'satisfaction_with_install_score', 'satisfied_with_training', 'satisfied_with_training_score', 'recommend_implementation_team', 'recommend_implementation_team_score', 'expectations_set', 'expectations_set_score', 'my_expectations_met', 'my_expectations_met_score', 'recommend_dt', 'recommend_dt_score', 'decision_making_yn', 'decision_making_yn_score', 'sales_based_met_expectations', 'sales_based_met_expectations_score', 'sales_satisfaction', 'sales_satisfaction_score', 'sales_rep_satisfaction', 'sales_rep_satisfaction_score', 'part_of_install', 'part_of_install_score', 'ease_of_use', 'ease_of_use_score', 'dms_effectiveness', 'dms_effectiveness_score', 'compelled', 'compelled_score', 'primary_role', 'primary_role_id', 'comments', 'unsatisfied_comments'];

} else if(fields_to_use_for_csv == 3) {
  // Month 3 Fields
  fields = ['response_number', 'submit_date', 'DISCARD_THIS_FIELD', 'go_live_date', 'dt_id', 'phone', 'dealership', 'respondent_name', 'primary_role', 'primary_role_id', 'satisfaction_with_dms', 'satisfaction_with_dms_score', 'satisfaction_with_billing_and_management_team', 'satisfaction_with_billing_and_management_team_score', 'satisfaction_with_support', 'satisfaction_with_support_score', 'initial_training_satisfaction', 'initial_training_satisfaction_score', 'recommend_dt', 'recommend_dt_score', 'unsatisfied_comments', 'comments'];

} else if(fields_to_use_for_csv == 6) {
  // Month 6 Fields
  fields = ['response_number', 'submit_date', 'DISCARD_THIS_FIELD', 'go_live_date', 'dt_id', 'phone', 'dealership', 'respondent_name', 'primary_role', 'primary_role_id', 'satisfaction_with_dms', 'satisfaction_with_dms_score', 'satisfaction_with_billing_and_management_team', 'satisfaction_with_billing_and_management_team_score', 'satisfaction_with_support', 'satisfaction_with_support_score', 'additional_training_received', 'additional_training_received_score', 'satisfaction_with_initial_training', 'satisfaction_with_initial_training_score', 'comfort_level_with_dms', 'comfort_level_with_dms_score', 'dms_effectiveness', 'dms_effectiveness_score', 'integration_rating', 'integration_rating_score', 'recommend_dt', 'recommend_dt_score', 'likelihood_of_repurchase', 'likelihood_of_repurchase_score', 'unsatisfied_comments', 'comments'];

} else if(fields_to_use_for_csv == 12) {
  // Month 12 Fields
  fields = ['response_number', 'submit_date', 'DISCARD_THIS_FIELD', 'go_live_date', 'dt_id', 'phone', 'dealership', 'respondent_name', 'primary_role', 'primary_role_id', 'satisfaction_with_dms', 'satisfaction_with_dms_score', 'satisfaction_with_billing_and_management_team', 'satisfaction_with_billing_and_management_team_score', 'satisfaction_with_support', 'satisfaction_with_support_score', 'additional_training_received', 'additional_training_received_score', 'additional_training_satisfaction', 'additional_training_satisfaction_score', 'comfort_level_with_dms', 'comfort_level_with_dms_score', 'dms_effectiveness', 'dms_effectiveness_score', 'integration_rating', 'integration_rating_score', 'recommend_dt', 'recommend_dt_score', 'likelihood_of_repurchase', 'likelihood_of_repurchase_score', 'unsatisfied_comments', 'comments'];

} else if(fields_to_use_for_csv == 18) {
  // Month 18 Fields
  fields = ['response_number', 'submit_date', 'DISCARD_THIS_FIELD', 'go_live_date', 'dt_id', 'DISCARD_THIS_FIELD', 'DISCARD_THIS_FIELD', 'DISCARD_THIS_FIELD', 'phone', 'dealership', 'respondent_name', 'primary_role', 'primary_role_id', 'satisfaction_with_dms', 'satisfaction_with_dms_score', 'satisfaction_with_billing_and_management_team', 'satisfaction_with_billing_and_management_team_score', 'satisfaction_with_support', 'satisfaction_with_support_score', 'additional_training_received', 'additional_training_received_score', 'comfort_level_with_dms', 'comfort_level_with_dms_score', 'dms_effectiveness', 'dms_effectiveness_score', 'integration_rating', 'integration_rating_score', 'recommend_dt', 'recommend_dt_score', 'recommend_explanation', 'likelihood_of_repurchase', 'likelihood_of_repurchase_score', 'unsatisfied_comments', 'comments'];

}

// xls to json library
var path_to_xls = './survey_files/' + process.env.FILE_TO_PARSE;

var json_objs = [];

if(typeof require !== 'undefined') XLS = require('xlsjs');
var workbook = XLS.readFile(path_to_xls);
var sheet = workbook.Sheets['Results'];

var convert_to_json = XLS.utils.sheet_to_json(sheet);

_.forEach(convert_to_json, function(a) {
  // need to remove "amp;", "<font color=red>", and "<font color=green>"

  var obj_string = JSON.stringify(a);
  var prep1 = obj_string.replace(/<font color=red>/g, '');
  var prep2 = prep1.replace(/<font color=green>/g, '');
  var prep3 = prep2.replace(/amp;/g, '');

  var b = JSON.parse(prep3);

  if(fields_to_use_for_csv == 1) {
    // Month 1 Fields
    var json_obj = {
      response_number: b.response_number,
      submit_date: b.submit_date,
      dt_id: b.dt_id,
      phone: b.phone,
      dealership: b.dealership,
      respondent_name: b.respondent_name,
      good_communication: b.good_communication,
      good_communication_score: b.good_communication_score,
      satisfaction_with_install: b.satisfaction_with_install,
      satisfaction_with_install_score: b.satisfaction_with_install_score,
      satisfied_with_training: b.satisfied_with_training,
      satisfied_with_training_score: b.satisfied_with_training_score,
      recommend_implementation_team: b.recommend_implementation_team,
      recommend_implementation_team_score: b.recommend_implementation_team_score,
      expectations_set: b.expectations_set,
      expectations_set_score: b.expectations_set_score,
      my_expectations_met: b.my_expectations_met,
      my_expectations_met_score: b.my_expectations_met_score,
      recommend_dt: b.recommend_dt,
      recommend_dt_score: b.recommend_dt_score,
      decision_making_yn: b.decision_making_yn,
      decision_making_yn_score: b.decision_making_yn_score,
      sales_based_met_expectations: b.sales_based_met_expectations,
      sales_based_met_expectations_score: b.sales_based_met_expectations_score,
      sales_satisfaction: b.sales_satisfaction,
      sales_satisfaction_score: b.sales_satisfaction_score,
      sales_rep_satisfaction: b.sales_rep_satisfaction,
      sales_rep_satisfaction_score: b.sales_rep_satisfaction_score,
      part_of_install: b.part_of_install,
      part_of_install_score: b.part_of_install_score,
      ease_of_use: b.ease_of_use,
      ease_of_use_score: b.ease_of_use_score,
      dms_effectiveness: b.dms_effectiveness,
      dms_effectiveness_score: b.dms_effectiveness_score,
      compelled: b.compelled,
      compelled_score: b.compelled_score,
      primary_role: b.primary_role,
      primary_role_id: b.primary_role_id,
      comments: b.comments,
      unsatisfied_comments: b.unsatisfied_comments
    };
  

  } else if(fields_to_use_for_csv == 3) {
    // Month 3 Fields
    var json_obj = {
      response_number: b.response_number,
      submit_date: b.submit_date,
      go_live_date: b.go_live_date,
      dt_id: b.dt_id,
      phone: b.phone,
      dealership: b.dealership,
      respondent_name: b.respondent_name,
      primary_role: b.primary_role,
      primary_role_id: b.primary_role_id,
      satisfaction_with_dms: b.satisfaction_with_dms,
      satisfaction_with_dms_score: b.satisfaction_with_dms_score,
      satisfaction_with_billing_and_management_team: b.satisfaction_with_billing_and_management_team,
      satisfaction_with_billing_and_management_team_score: b.satisfaction_with_billing_and_management_team_score,
      satisfaction_with_support: b.satisfaction_with_support,
      satisfaction_with_support_score: b.satisfaction_with_support_score,
      initial_training_satisfaction: b.initial_training_satisfaction,
      initial_training_satisfaction_score: b.initial_training_satisfaction_score,
      recommend_dt: b.recommend_dt,
      recommend_dt_score: b.recommend_dt_score,
      unsatisfied_comments: b.unsatisfied_comments,
      comments: b.comments
    };
  
  } else if(fields_to_use_for_csv == 6) {
    // Month 6 Fields
    var json_obj = {
      response_number: b.response_number,
      submit_date: b.submit_date,
      go_live_date: b.go_live_date,
      dt_id: b.dt_id,
      phone: b.phone,
      dealership: b.dealership,
      respondent_name: b.respondent_name,
      primary_role: b.primary_role,
      primary_role_id: b.primary_role_id,
      satisfaction_with_dms: b.satisfaction_with_dms,
      satisfaction_with_dms_score: b.satisfaction_with_dms_score,
      satisfaction_with_billing_and_management_team: b.satisfaction_with_billing_and_management_team,
      satisfaction_with_billing_and_management_team_score: b.satisfaction_with_billing_and_management_team_score,
      satisfaction_with_support: b.satisfaction_with_support,
      satisfaction_with_support_score: b.satisfaction_with_support_score,
      additional_training_received: b.additional_training_received,
      additional_training_received_score: b.additional_training_received_score,
      satisfaction_with_initial_training: b.satisfaction_with_initial_training,
      satisfaction_with_initial_training_score: b.satisfaction_with_initial_training_score,
      comfort_level_with_dms: b.comfort_level_with_dms,
      comfort_level_with_dms_score: b.comfort_level_with_dms_score,
      dms_effectiveness: b.dms_effectiveness,
      dms_effectiveness_score: b.dms_effectiveness_score,
      integration_rating: b.integration_rating,
      integration_rating_score: b.integration_rating_score,
      recommend_dt: b.recommend_dt,
      recommend_dt_score: b.recommend_dt_score,
      likelihood_of_repurchase: b.likelihood_of_repurchase,
      likelihood_of_repurchase_score: b.likelihood_of_repurchase_score,
      unsatisfied_comments: b.unsatisfied_comments,
      comments: b.comments
    };
  

  } else if(fields_to_use_for_csv == 12) {
    // Month 12 Fields
    var json_obj = {
      response_number: b.response_number,
      submit_date: b.submit_date,
      go_live_date: b.go_live_date,
      dt_id: b.dt_id,
      phone: b.phone,
      dealership: b.dealership,
      respondent_name: b.respondent_name,
      primary_role: b.primary_role,
      primary_role_id: b.primary_role_id,
      satisfaction_with_dms: b.satisfaction_with_dms,
      satisfaction_with_dms_score: b.satisfaction_with_dms_score,
      satisfaction_with_billing_and_management_team: b.satisfaction_with_billing_and_management_team,
      satisfaction_with_billing_and_management_team_score: b.satisfaction_with_billing_and_management_team_score,
      satisfaction_with_support: b.satisfaction_with_support,
      satisfaction_with_support_score: b.satisfaction_with_support_score,
      additional_training_received: b.additional_training_received,
      additional_training_received_score: b.additional_training_received_score,
      additional_training_satisfaction: b.additional_training_satisfaction,
      additional_training_satisfaction_score: b.additional_training_satisfaction_score,
      comfort_level_with_dms: b.comfort_level_with_dms,
      comfort_level_with_dms_score: b.comfort_level_with_dms_score,
      dms_effectiveness: b.dms_effectiveness,
      dms_effectiveness_score: b.dms_effectiveness_score,
      integration_rating: b.integration_rating,
      integration_rating_score: b.integration_rating_score,
      recommend_dt: b.recommend_dt,
      recommend_dt_score: b.recommend_dt_score,
      likelihood_of_repurchase: b.likelihood_of_repurchase,
      likelihood_of_repurchase_score: b.likelihood_of_repurchase_score,
      unsatisfied_comments: b.unsatisfied_comments,
      comments: b.comments
    };
  

  } else if(fields_to_use_for_csv == 18) {
    // Month 18 Fields
    var json_obj = {
      response_number: b.response_number, 
      submit_date: b.submit_date, 
      DISCARD_THIS_FIELD: b.DISCARD_THIS_FIELD, 
      go_live_date: b.go_live_date, 
      dt_id: b.dt_id, 
      phone: b.phone, 
      dealership: b.dealership, 
      respondent_name: b.respondent_name, 
      primary_role: b.primary_role, 
      primary_role_id: b.primary_role_id, 
      satisfaction_with_dms: b.satisfaction_with_dms, 
      satisfaction_with_dms_score: b.satisfaction_with_dms_score, 
      satisfaction_with_billing_and_management_team: b.satisfaction_with_billing_and_management_team, 
      satisfaction_with_billing_and_management_team_score: b.satisfaction_with_billing_and_management_team_score, 
      satisfaction_with_support: b.satisfaction_with_support, 
      satisfaction_with_support_score: b.satisfaction_with_support_score, 
      additional_training_received: b.additional_training_received, 
      additional_training_received_score: b.additional_training_received_score, 
      comfort_level_with_dms: b.comfort_level_with_dms, 
      comfort_level_with_dms_score: b.comfort_level_with_dms_score, 
      dms_effectiveness: b.dms_effectiveness, 
      dms_effectiveness_score: b.dms_effectiveness_score, 
      integration_rating: b.integration_rating, 
      integration_rating_score: b.integration_rating_score, 
      recommend_dt: b.recommend_dt, 
      recommend_dt_score: b.recommend_dt_score, 
      recommend_explanation: b.recommend_explanation, 
      likelihood_of_repurchase: b.likelihood_of_repurchase, 
      likelihood_of_repurchase_score: b.likelihood_of_repurchase_score, 
      unsatisfied_comments: b.unsatisfied_comments, 
      comments: b.comments
    };
  

  }

  console.log(json_obj);
  json_objs.push(json_obj);

});

json2csv({ data: json_objs, fields: fields }, function(err, csv) {
  if (err) console.log(err);
  fs.writeFile(filename, csv, function(err) {
    if (err) throw err;
    console.log('file saved');
  });
});

// console.log(convert_to_json);

module.exports = app;