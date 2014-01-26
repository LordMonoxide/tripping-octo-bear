<?php

Route::group(['prefix' => 'api'], function() {
  Route::group(['prefix' => 'auth'], function() {
    Route::get ('/check',    ['as' => 'api.auth.check',    'uses' => 'api\AuthController@check']);
    Route::put ('/register', ['as' => 'api.auth.register', 'uses' => 'api\AuthController@register']);
    Route::post('/login',    ['as' => 'api.auth.login',    'uses' => 'api\AuthController@login']);
    Route::post('/logout',   ['as' => 'api.auth.logout',   'uses' => 'api\AuthController@logout']);
    Route::get ('/security', ['as' => 'api.auth.security', 'uses' => 'api\AuthController@security']);
    Route::post('/security', ['as' => 'api.auth.security', 'uses' => 'api\AuthController@unlock']);
  });
  
  Route::group(['prefix' => 'storage'], function() {
    Route::group(['prefix' => 'characters'], function() {
      Route::get   ('/', ['as' => 'api.storage.characters', 'uses' => 'api\storage\CharacterController@all']);
      Route::put   ('/', ['as' => 'api.storage.characters', 'uses' => 'api\storage\CharacterController@create']);
      Route::delete('/', ['as' => 'api.storage.characters', 'uses' => 'api\storage\CharacterController@delete']);
    });
  });
});