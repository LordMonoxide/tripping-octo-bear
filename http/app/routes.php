<?php

Route::group(['prefix' => 'api'], function() {
  Route::group(['prefix' => 'auth'], function() {
    Route::get('/check',    ['as' => 'api.auth.check',    'uses' => 'api\AuthController@check']);
    Route::put('/register', ['as' => 'api.auth.register', 'uses' => 'api\AuthController@register']);
    Route::post('/login',    ['as' => 'api.auth.login',    'uses' => 'api\AuthController@login']);
    Route::post('/logout',   ['as' => 'api.auth.logout',   'uses' => 'api\AuthController@logout']);
  });
  
  Route::group(['prefix' => 'storage'], function() {
    Route::get('/characters', ['as' => 'api.storage.characters', 'uses' => 'api\StorageController@characters']);
  });
});