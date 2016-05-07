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

    companyApi.findForEmail("steve@test.com").then(function(company) {
      vm.title = company;
    });
  }

})();
