(function(){
  'use strict';

  angular.module('officeAddin')
         .controller('homeController', ['companyApi', homeController]);

  /**
   * Controller constructor
   */
  function homeController(companyApi){
    var vm = this;  // jshint ignore:line
    vm.title = 'home controller';

    companyApi.findForEmail("steve@testi.com").then(function(company) {
      vm.user_name = "William Masterson";
      vm.company_name = "Mulesoft";
      vm.mrr = "$1200";
      vm.licenses = "20"
      vm.client_age = "1 Year, 3 Months";
      vm.account_manager = "Jessica Ryan";
      vm.location = "San Francisco, CA";
    });
  }

})();
