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
      }

    return new CompanyApi();
  }]);
})();
