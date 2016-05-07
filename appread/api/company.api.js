(function(){
  'use strict';

  angular.module('officeAddin')
         .service('companyApi', ["$q", function($q) {
           class CompanyApi {
             findForEmail(email) {
               return $q(function(resolve) {
                 resolve(email);
               });
             }
           }

           return new CompanyApi();
         }]);
})();
