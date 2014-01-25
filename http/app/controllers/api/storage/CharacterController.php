<?php namespace api\storage;

use Auth;
use Controller;
use Input;
use Response;
use Validator;

use Character;

class CharacterController extends Controller {
  public function __construct() {
    $this->beforeFilter('auth.401');
  }
  
  public function all() {
    return Response::json(Auth::user()->characters, 200);
  }
  
  public function create() {
    $validator = Validator::make(Input::all(), [
      'name' => ['required', 'min:4', 'max:20', 'unique:characters,name'],
      'sex'  => ['required', 'in:male,female']
    ]);
    
    if($validator->passes()) {
      $char = new Character;
      $char->user()->associate(Auth::user());
      $char->name = Input::get('name');
      $char->sex  = Input::get('sex');
      $char->map  = 0;
      $char->x    = 0;
      $char->y    = 0;
      $char->dir  = 'down';
      $char->save();
      
      return Response::json(null, 201);
    } else {
      return Response::json($validator->messages(), 409);
    }
  }
}