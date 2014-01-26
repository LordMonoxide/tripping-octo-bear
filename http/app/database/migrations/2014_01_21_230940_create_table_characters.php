<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableCharacters extends Migration {
  public function up() {
    Schema::create('characters', function($table) {
      $table->increments('id');
      $table->integer('user_id')->unsigned();
      $table->integer('auth_id')->unsigned()->nullable();
      $table->integer('guild_id')->unsigned()->nullable();
      
      $table->string('name', 20);
      $table->enum('sex', ['male', 'female']);
      
      $table->integer('lvl')->unsigned()->default(1);
      $table->integer('exp')->unsigned()->default(0);
      $table->integer('pts')->unsigned()->default(0);
      
      $table->integer('hp')->unsigned()->default(0);
      $table->integer('mp')->unsigned()->default(0);
      
      $table->integer('str')->unsigned()->default(1);
      $table->integer('end')->unsigned()->default(1);
      $table->integer('int')->unsigned()->default(1);
      $table->integer('agl')->unsigned()->default(1);
      $table->integer('wil')->unsigned()->default(1);
      
      $table->integer('weapon')->unsigned()->nullable();
      $table->integer('armour')->unsigned()->nullable();
      $table->integer('shield')->unsigned()->nullable();
      $table->integer('aura')->unsigned()->nullable();
      
      $table->integer('clothes')->unsigned()->nullable();
      $table->integer('gear')->unsigned()->nullable();
      $table->integer('hair')->unsigned()->nullable();
      $table->integer('head')->unsigned()->nullable();
      
      $table->integer('map')->unsigned();
      $table->integer('x')->unsigned();
      $table->integer('y')->unsigned();
      $table->enum('dir', ['up', 'down', 'left', 'right', 'upleft', 'upright', 'downleft', 'downright']);
      
      $table->boolean('threshold')->default(false);
      
      $table->timestamps();
      
      $table->foreign('user_id')
             ->references('id')
             ->on('users');
      
      $table->foreign('guild_id')
             ->references('id')
             ->on('guilds');
    });
  }
  
  public function down() {
    Schema::drop('characters');
  }
}