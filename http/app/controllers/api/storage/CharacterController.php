<?php namespace api\storage;

use Auth;
use Controller;
use Input;
use Request;
use Response;
use Validator;

use Character;

class CharacterController extends Controller {
  public function __construct() {
    $this->beforeFilter('user.security');
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
      $char->map  = 1;
      $char->x    = 0;
      $char->y    = 0;
      $char->dir  = 'down';
      $char->save();
      
      return Response::json(null, 201);
    } else {
      return Response::json($validator->messages(), 409);
    }
  }
  
  public function delete() {
    $validator = Validator::make(Input::all(), [
      'id' => ['required', 'integer', 'exists:characters,id']
    ]);
    
    if($validator->passes()) {
      Character::destroy(Input::get('id'));
      return Response::json(null, 200);
    } else {
      return Response::json($validator->messages(), 409);
    }
  }
  
  public function choose() {
    $validator = Validator::make(Input::all(), [
      'id' => ['required', 'integer', 'exists:characters,id,user_id,' . Auth::user()->id]
    ]);
    
    if($validator->passes()) {
      $ip = Auth::user()->ips()->where('ip', '=', ip2long(Request::getClientIp()))->first();
      
      $char = Character::find(Input::get('id'));
      $char->auth()->associate($ip);
      $char->save();
      
      return Response::json(['u_id' => Auth::user()->id, 'c_id' => $char->id], 200);
    } else {
      return Response::json($validator->messages(), 409);
    }
  }
}