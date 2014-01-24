<?php

use Illuminate\Database\Migrations\Migration;

class CreateTableCharacterItems extends Migration {
  public function up() {
    Schema::create('character_items', function($table) {
      $table->increments('id');
      $table->integer('character_id')->unsigned();
      $table->integer('item_id')->unsigned();
      
      $table->integer('value')->unsigned();
      $table->boolean('bound');
      
      $table->foreign('character_id')
             ->references('id')
             ->on('characters');
      
      $table->foreign('item_id')
             ->references('id')
             ->on('items');
    });
  }
  
  public function down() {
    Schema::drop('character_items');
  }
}