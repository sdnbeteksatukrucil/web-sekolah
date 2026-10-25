-- Jalankan seluruh isi file ini di Supabase > SQL Editor
create table profiles(id uuid primary key references auth.users on delete cascade, role text not null check(role in('admin','guru','siswa')), nama text, kelas int, siswa_id uuid, guru_id uuid);
create table siswa(id uuid primary key default gen_random_uuid(), nama text not null, nisn text, nis text, nik text, ttl text, kelas int not null check(kelas between 1 and 6), ayah text, ibu text, pekerjaan text, hp text, created_at timestamptz default now());
create table guru(id uuid primary key default gen_random_uuid(), nama text not null, nip text, kelas int, hp text, drive_url text, created_at timestamptz default now());
create table absensi(id uuid primary key default gen_random_uuid(), siswa_id uuid not null references siswa on delete cascade, tanggal date not null, status text not null check(status in('H','S','I','A')), unique(siswa_id,tanggal));
create table nilai(id uuid primary key default gen_random_uuid(), siswa_id uuid not null references siswa on delete cascade, mapel text not null, semester text not null, nilai numeric not null, created_at timestamptz default now());
create table catatan(id uuid primary key default gen_random_uuid(), siswa_id uuid not null references siswa on delete cascade, jenis text not null check(jenis in('penghargaan','pelanggaran')), poin int not null default 0, deskripsi text, tanggal date default current_date, created_at timestamptz default now());
create table pengaturan(key text primary key, data jsonb not null default '{}');
create table berita(id uuid primary key default gen_random_uuid(), jenis text not null default 'berita' check(jenis in('berita','pengumuman')), judul text not null, isi text, gambar text, tanggal date default current_date, created_at timestamptz default now());
create table galeri(id uuid primary key default gen_random_uuid(), judul text, url text not null, created_at timestamptz default now());
create index on siswa(kelas); create index on absensi(tanggal); create index on absensi(siswa_id); create index on nilai(siswa_id); create index on catatan(siswa_id);

-- fungsi bantu (security definer agar tidak rekursif)
create function my_role() returns text language sql stable security definer set search_path=public as $$ select role from profiles where id=auth.uid() $$;
create function my_kelas() returns int language sql stable security definer set search_path=public as $$ select kelas from profiles where id=auth.uid() $$;
create function my_siswa() returns uuid language sql stable security definer set search_path=public as $$ select siswa_id from profiles where id=auth.uid() $$;
create function my_guru() returns uuid language sql stable security definer set search_path=public as $$ select guru_id from profiles where id=auth.uid() $$;
create function kelas_siswa(sid uuid) returns int language sql stable security definer set search_path=public as $$ select kelas from siswa where id=sid $$;

alter table profiles enable row level security; alter table siswa enable row level security; alter table guru enable row level security;
alter table pengaturan enable row level security; alter table berita enable row level security; alter table galeri enable row level security;

create policy p_sel on profiles for select using(id=auth.uid() or my_role()='admin');
create policy s_sel on siswa for select using(my_role()='admin' or (my_role()='guru' and kelas=my_kelas()) or (my_role()='siswa' and id=my_siswa()));
create policy s_w on siswa for all using(my_role()='admin') with check(my_role()='admin');
create policy g_sel on guru for select using(my_role()='admin' or id=my_guru());
create policy g_w on guru for all using(my_role()='admin') with check(my_role()='admin');

-- absensi, nilai, catatan: admin semua, guru kelasnya, siswa miliknya (baca saja)
do $$ declare t text; begin foreach t in array array['absensi','nilai','catatan'] loop
  execute format('alter table %I enable row level security',t);
  execute format('create policy r_%1$s on %1$I for select using(my_role()=''admin'' or (my_role()=''guru'' and kelas_siswa(siswa_id)=my_kelas()) or (my_role()=''siswa'' and siswa_id=my_siswa()))',t);
  execute format('create policy w_%1$s on %1$I for all using(my_role()=''admin'' or (my_role()=''guru'' and kelas_siswa(siswa_id)=my_kelas())) with check(my_role()=''admin'' or (my_role()=''guru'' and kelas_siswa(siswa_id)=my_kelas()))',t);
end loop; end $$;

-- konten publik: semua orang boleh baca, hanya admin yang boleh ubah
do $$ declare t text; begin foreach t in array array['pengaturan','berita','galeri'] loop
  execute format('create policy pub_r_%1$s on %1$I for select using(true)',t);
  execute format('create policy pub_w_%1$s on %1$I for all using(my_role()=''admin'') with check(my_role()=''admin'')',t);
end loop; end $$;

-- penyimpanan foto/logo
insert into storage.buckets(id,name,public) values('media','media',true) on conflict do nothing;
create policy media_r on storage.objects for select using(bucket_id='media');
create policy media_w on storage.objects for all to authenticated using(bucket_id='media' and my_role()='admin') with check(bucket_id='media' and my_role()='admin');
