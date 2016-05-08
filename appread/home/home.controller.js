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
    companyApi.findForEmail(vm.email_address).then(function(resp) {
      vm.company_html = enrichCompanyHtml(resp.data.company_html);
      vm.user_html = resp.data.html;
    }).catch(function(err) {
      return companyApi.findCompanyForEmail(vm.email_address).then(function(resp) {
        vm.company_html = enrichCompanyHtml(resp.data.company_html);
      }).catch(function(err) {
        vm.not_found = true;
      });
    }).finally(function() {
      vm.loading = false;
    });
  }

  function enrichCompanyHtml(html) {
    var el = angular.element(html);
    el.find("h3").prepend(angular.element("<i class='ms-Icon ms-Icon--work'>"))
    return el.html();
  }

})();
