<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableCharacters extends Migration {
  public function up() {
    Schema::create('characters', function($table) {
      $table->increments('id');
      $table->integer('user_id')->unsigned();
      $table->integer('guild_id')->unsigned();
      
      $table->string('name', 20);
      $table->enum('sex', ['male', 'female']);
      
      $table->tinyInteger('lvl');
      $table->integer('exp');
      $table->integer('pts');
      
      $table->integer('hp');
      $table->integer('mp');
      
      $table->integer('str');
      $table->integer('end');
      $table->integer('int');
      $table->integer('agl');
      $table->integer('wil');
      
      $table->integer('weapon');
      $table->integer('armour');
      $table->integer('shield');
      $table->integer('aura');
      
      $table->integer('clothes');
      $table->integer('gear');
      $table->integer('hair');
      $table->integer('head');
      
      $table->integer('map');
      $table->tinyInteger('x');
      $table->tinyInteger('y');
      $table->tinyInteger('dir');
      
      $table->tinyInteger('threshold');
      
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