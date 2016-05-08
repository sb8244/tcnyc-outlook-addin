(function(){
  'use strict';

  const URL = 'https://tcnyc.herokuapp.com';
  //const URL = 'https://stevegrok.ngrok.io';
  angular.module('officeAddin').service('companyApi', ["$http", "$q", function($http, $q) {
    class CompanyApi {
      findForEmail(email) {
        return $http({
          method: 'GET',
          url: URL + '/api/users/find?email=' + email
        });
      }

      findCompanyForEmail(email) {
        if (email) {
          var domain = email.split("@").splice(-1)[0];
          return $http({
            method: 'GET',
            url: URL + '/api/companies/find?domain=' + domain
          });
        } else {
          return $q(function(resolve, reject) {
            reject();
          });
        }
      }
    }

    return new CompanyApi();
  }]);
})();
