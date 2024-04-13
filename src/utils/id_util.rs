use std::collections::hash_map::DefaultHasher;
use std::collections::HashSet;
use std::hash::{Hash, Hasher};
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{SystemTime, UNIX_EPOCH};

pub(crate) fn new_id() -> u64 {
    let time = SystemTime::now().duration_since(UNIX_EPOCH).unwrap().as_micros();
    let mut hasher = DefaultHasher::new();
    time.hash(&mut hasher);
    std::process::id().hash(&mut hasher);
    thread::current().id().hash(&mut hasher);
    hasher.finish()
}

#[test]
fn test_id() {
    let mut v = Arc::new(Mutex::new(vec![]));
    for _ in 0..10 {
        let handles: Vec<_> = (0..5000).map(|i| {
            let atomic_vec = Arc::clone(&v);
            thread::spawn(move || {
                atomic_vec.lock().unwrap().push(new_id());
            })
        }).collect();
        for handle in handles {
            handle.join().unwrap();
        }
    }
    let len_before = v.lock().unwrap().len();
    let len: usize = v.lock().unwrap().iter().collect::<HashSet<&u64>>().iter().count();
    assert_eq!(len_before, len);
}