<?php

class DatabaseSeeder extends Seeder {
  public function run() {
    Eloquent::unguard();
    $this->call('TableTruncater');
    $this->call('UserTableSeeder');
    $this->call('ItemTableSeeder');
    $this->call('NPCTableSeeder');
  }
}

class UserTableSeeder extends Seeder {
  public function run() {
    $user = User::create([
      'email'      => 'corey@narwhunderful.com',
      'password'   => Hash::make('monoxide'),
      'name_first' => 'Corey',
      'name_last'  => 'Frenette'
    ]);
    
    UserSecurityQuestion::create([
      'user_id' => $user->id,
      'question' => 'The answer to this question is 1',
      'answer' => '1'
    ]);
    
    UserSecurityQuestion::create([
      'user_id' => $user->id,
      'question' => 'The answer to this question is 2',
      'answer' => '2'
    ]);
    
    UserSecurityQuestion::create([
      'user_id' => $user->id,
      'question' => 'The answer to this question is 3',
      'answer' => '3'
    ]);
  }
}

class ItemTableSeeder extends Seeder {
  public function run() {
    Item::create([
      'name'       => 'Sword',
      'desc'       => 'A basic weapon',
      'sound'      => '',
      'pic'        => 1,
      'type'       => 1,
      'data1'      => 10,
      'data2'      => 0,
      'data3'      => 0,
      'access_req' => 0,
      'level_req'  => 0,
      'price'      => 10,
      'rarity'     => 0,
      'speed'      => 0,
      'bind_type'  => 0,
      'animation'  => 0,
      'add_hp'     => 0,
      'add_mp'     => 0,
      'add_exp'    => 0,
      'add_str'    => 0,
      'add_end'    => 0,
      'add_int'    => 0,
      'add_agl'    => 0,
      'add_wil'    => 0,
      'req_str'    => 0,
      'req_end'    => 0,
      'req_int'    => 0,
      'req_agl'    => 0,
      'req_wil'    => 0,
      'aura'       => 0,
      'projectile' => 0,
      'range'      => 0,
      'rotation'   => 0,
      'ammo'       => 0,
      'two_handed' => 0,
      'stackable'  => 0,
      'p_def'      => 0,
      'r_def'      => 0,
      'm_def'      => 0,
      'container_0'        => 0,
      'container_chance_0' => 0,
      'container_1'        => 0,
      'container_chance_1' => 0,
      'container_2'        => 0,
      'container_chance_2' => 0,
      'container_3'        => 0,
      'container_chance_3' => 0,
      'container_4'        => 0,
      'container_chance_4' => 0
    ]);
  }
}

class NPCTableSeeder extends Seeder {
  public function run() {
    NPC::create([
      'name' => 'Test NPC',
      'say' => 'Blarg',
      'sound' => '',
      
      'sprite' => 1,
      'spawn_secs' => 3,
      'behaviour' => 0,
      'range' => 10,
      
      'lvl' => 1,
      'exp' => 10,
      'exp_max' => 100,
      'hp' => 50,
      'str' => 2,
      'end' => 2,
      'int' => 2,
      'agl' => 2,
      'wil' => 2,
      
      'animation' => 0,
      'damage' => 2,
      'quest' => 0,
      'quest_num' => 0,
      
      'event' => 0,
      
      'projectile' => 0,
      'projectile_range' => 0,
      'rotation' => 0,
      'moral' => 0,
      
      'colour' => 0,
      
      'spawn_at_day' => true,
      'spawn_at_night' => false
    ]);
  }
}

class TableTruncater extends Seeder {
  public function run() {
    $this->command->info('Getting foreign keys...');
    $t1 = microtime(true);
    
    // Get the database name
    $dbname = DB::connection('mysql')->getDatabaseName();
    
    // Find the FKs
    $fks = DB::table('INFORMATION_SCHEMA.KEY_COLUMN_USAGE')
            ->select('TABLE_NAME', 'COLUMN_NAME', 'CONSTRAINT_NAME', 'REFERENCED_TABLE_NAME', 'REFERENCED_COLUMN_NAME')
      ->whereNotNull('REFERENCED_TABLE_NAME')
               ->get();
    
    // Find the tables
    $tables = DB::table('INFORMATION_SCHEMA.TABLES')
               ->select('TABLE_SCHEMA', 'TABLE_NAME')
                ->where('TABLE_SCHEMA', '=', $dbname)
                ->where('TABLE_NAME', '<>', 'migrations')
                  ->get();
    
    $this->command->info('Killing foreign keys...');
    
    // Kill all FKs
    foreach($fks as $fk) {
      Schema::table($fk->TABLE_NAME, function($table) use($fk) {
        $table->dropForeign($fk->CONSTRAINT_NAME);
      });
    }
    
    $this->command->info('Truncating tables...');
    
    // Truncate all tables
    foreach($tables as $table) {
      DB::table($table->TABLE_NAME)->truncate();
    }
    
    $this->command->info('Reinstating foreign keys...');
    
    // Add all the FKs back
    foreach($fks as $fk) {
      Schema::table($fk->TABLE_NAME, function($table) use($fk) {
        $table->foreign($fk->COLUMN_NAME)
              ->references($fk->REFERENCED_COLUMN_NAME)
              ->on($fk->REFERENCED_TABLE_NAME);
      });
    }
    
    $this->command->info('Truncation completed in ' . (microtime(true) - $t1) . ' seconds.');
  }
}