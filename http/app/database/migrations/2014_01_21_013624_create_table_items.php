<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableItems extends Migration {
  public function up() {
    Schema::create('items', function($table) {
      $table->increments('id');
      $table->string('name', 40);
      $table->string('desc', 256);
      $table->string('sound', 40);
      $table->integer('pic')->unsigned();
      $table->integer('type')->unsigned();
      $table->integer('data1')->unsigned();
      $table->integer('data2')->unsigned();
      $table->integer('data3')->unsigned();
      $table->integer('access_req')->unsigned();
      $table->integer('level_req')->unsigned();
      $table->integer('price')->unsigned();
      $table->integer('rarity')->unsigned();
      $table->integer('speed')->unsigned();
      $table->integer('bind_type')->unsigned();
      $table->integer('animation')->unsigned();
      $table->integer('add_hp')->unsigned();
      $table->integer('add_mp')->unsigned();
      $table->integer('add_exp')->unsigned();
      $table->integer('add_str')->unsigned();
      $table->integer('add_end')->unsigned();
      $table->integer('add_int')->unsigned();
      $table->integer('add_agl')->unsigned();
      $table->integer('add_wil')->unsigned();
      $table->integer('req_str')->unsigned();
      $table->integer('req_end')->unsigned();
      $table->integer('req_int')->unsigned();
      $table->integer('req_agl')->unsigned();
      $table->integer('req_wil')->unsigned();
      $table->integer('aura')->unsigned();
      $table->integer('projectile')->unsigned();
      $table->integer('range')->unsigned();
      $table->integer('rotation')->unsigned();
      $table->integer('ammo')->unsigned();
      $table->integer('two_handed')->unsigned();
      $table->integer('stackable')->unsigned();
      $table->integer('p_def')->unsigned();
      $table->integer('r_def')->unsigned();
      $table->integer('m_def')->unsigned();
      
      for($i = 0; $i < 5; $i++) {
        $table->integer('container_' . $i)->unsigned();
        $table->integer('container_chance_' . $i)->unsigned();
      }
      
      $table->timestamps();
    });
  }
  
  public function down() {
    Schema::drop('items');
  }
}