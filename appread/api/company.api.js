(function(){
  'use strict';

  angular.module('officeAddin').service('companyApi', ["$http", function($http) {
    class CompanyApi {
      findForEmail(email) {
        return $http({
          method: 'GET',
          url: 'https://tcnyc.herokuapp.com/api/users/find?email=' + email
        });
      }

      findCompanyForEmail(email) {
        var domain = email.split("@").splice(-1)[0];
        return $http({
          method: 'GET',
          url: 'https://tcnyc.herokuapp.com/api/companies/find?domain=' + domain
        });
      }
    }

    return new CompanyApi();
  }]);
})();
