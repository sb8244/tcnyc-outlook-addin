(function(){
  'use strict';

  angular.module('officeAddin')
         .controller('homeController', ['companyApi', homeController]);

  /**
   * Controller constructor
   */
  function homeController(companyApi){
    var vm = this;  // jshint ignore:line
    vm.loading = true;

    vm.email_address = Office.context.mailbox.item.from.emailAddress;
    companyApi.findForEmail(Office.context.mailbox.item.from.emailAddress).then(function(resp) {
      vm.company_html = resp.data.company_html;
      vm.user_html = resp.data.user_html;
    }).catch(function(err) {
      vm.not_found = true;
    }).finally(function() {
      vm.loading = false;
    });
  }

})();
