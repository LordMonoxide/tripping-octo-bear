<?php namespace api;

use Controller;

class StorageController extends Controller {
  public function __construct() {
    $this->beforeFilter('auth.401');
  }
  
  public function characters() {
    return Response::json(Auth::user()->characters, 200);
  }
}