<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableNpcs extends Migration {
  public function up() {
    Schema::create('npcs', function($table) {
      $table->increments('id');
      $table->string('name', 40);
      $table->string('say', 256);
      $table->string('sound', 40);
      
      $table->integer('sprite')->unsigned();
      $table->integer('spawn_secs')->unsigned();
      $table->tinyInteger('behaviour')->unsigned();
      $table->tinyInteger('range')->unsigned();
      
      $table->integer('lvl')->unsigned();
      $table->integer('exp')->unsigned();
      $table->integer('exp_max')->unsigned();
      $table->integer('hp')->unsigned();
      $table->integer('str')->unsigned();
      $table->integer('end')->unsigned();
      $table->integer('int')->unsigned();
      $table->integer('agl')->unsigned();
      $table->integer('wil')->unsigned();
      
      $table->integer('animation')->unsigned();
      $table->integer('damage')->unsigned();
      $table->tinyInteger('quest')->unsigned();
      $table->integer('quest_num')->unsigned();
      
      $table->integer('event')->unsigned();
      
      $table->integer('projectile')->unsigned();
      $table->tinyInteger('projectile_range')->unsigned();
      $table->smallInteger('rotation')->unsigned();
      $table->tinyInteger('moral')->unsigned();
      
      $table->integer('colour')->unsigned();
      
      $table->boolean('spawn_at_day');
      $table->boolean('spawn_at_night');
      
      $table->timestamps();
    });
  }
  
  public function down() {
    Schema::drop('npcs');
  }
}