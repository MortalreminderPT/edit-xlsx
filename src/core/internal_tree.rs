use std::cell::RefCell;
use std::collections::HashSet;
use std::fmt::Display;
use std::rc::Rc;
// trait Value {
//     // fn display(&self) -> String;
// }
//
// impl Value for String {
//     // fn display(&self) -> String {
//     //     self.to_string()
//     // }
// }
//
// impl Value for &str {
//     // fn display(&self) -> String {
//     //     self.to_string()
//     // }
// }

struct Node<T: Clone + Default> {
    left: i32,
    right: i32,
    value: T,
}

impl<T: Clone + Default> Node<T> {
    fn new(left: i32, right: i32, value: &T) -> Node<T> {
        Node {
            left,
            right,
            value: value.clone(),
        }
    }
}

pub(crate) struct InternalTree<T: Clone + Default> {
    node: Rc<RefCell<Node<T>>>,
    left_child: Option<Rc<RefCell<InternalTree<T>>>>,
    right_child: Option<Rc<RefCell<InternalTree<T>>>>,
}

impl<T: Clone + Default> InternalTree<T> {
    fn new() -> InternalTree<T> {
        InternalTree {
            node: Rc::new(RefCell::new(Node::new(0, 1, &Default::default()))),
            left_child: None,
            right_child: None,
        }
    }
    fn from_internal(left: i32, right: i32) -> InternalTree<T> {
        InternalTree {
            node: Rc::new(RefCell::new(Node::new(left, right, &Default::default()))),
            left_child: None,
            right_child: None,
        }
    }
    fn from_node(node: Node<T>) -> InternalTree<T> {
        let node = Rc::new(RefCell::new(node));
        InternalTree {
            node,
            left_child: None,
            right_child: None,
        }
    }
    fn update_node(&mut self, new_left: i32, new_right: i32, value: &T) {
        if new_left >= new_right { return }
        let node = &self.node;
        let left = node.borrow().left;
        let right = node.borrow().right;
        let old_value = &self.node.borrow().value.clone();
        if left >= new_left && right <= new_right {
            self.node.borrow_mut().value = value.clone();
        }
        if left == new_left && right == new_right {
            return;
        }
        // println!("Adding [{new_left}, {new_right}], {old_value} -> {}", value.display());
        if new_right <= left {
            let node = Node::new(new_left, new_right, value);
            let tree = InternalTree::from_node(node);
            match &self.left_child {
                None => self.left_child = Some(Rc::new(RefCell::new(tree))),
                Some(left_child) => {
                    left_child.borrow_mut().update_node(new_left, new_right, value);
                },
            };
        } else if new_left >= right {
            let node = Node::new(new_left, new_right, value);
            let tree = InternalTree::from_node(node);
            match &self.right_child {
                None => self.right_child = Some(Rc::new(RefCell::new(tree))),
                Some(right_child) => right_child.borrow_mut().update_node(new_left, new_right, value),
            };
        } else {
            let mut set = HashSet::new();
            set.insert(left);
            set.insert(right);
            set.insert(new_left);
            set.insert(new_right);
            let mut inters = set.iter().map(|v| *v).collect::<Vec<i32>>();// vec![left, right, new_left, new_right];
            inters.sort();
            let mut lid = 0;
            for i in 0..inters.len() - 1 {
                if inters[i] == left {
                    self.node.borrow_mut().right = inters[i + 1];
                    if inters[i] >= new_left && inters[i + 1] <= new_right {
                        self.node.borrow_mut().value = value.clone();
                    }
                    lid = i;
                    break;
                }
            }
            for i in 0..inters.len() - 1 {
                if i != lid {
                    if inters[i] >= new_left && inters[i + 1] <= new_right {
                        self.update_node(inters[i], inters[i + 1], value);
                    } else {
                        self.update_node(inters[i], inters[i + 1], old_value);
                    }
                }
            }
        }
    }
    // fn display(&self) {
    //     match &self.left_child {
    //         None => {},
    //         Some(left_child) => left_child.borrow().display(),
    //     };
    //     // println!("[{}, {}]: {}", self.node.borrow().left, self.node.borrow().right, self.node.borrow().value.display());
    //     match &self.right_child {
    //         None => {},
    //         Some(right_child) => right_child.borrow().display(),
    //     };
    // }
    fn recurse_insert(&self, v: &mut Vec<(i32, i32, T)>) {
        match &self.left_child {
            None => {},
            Some(left_child) => left_child.borrow().recurse_insert(v),
        };
        v.push((self.node.borrow().left, self.node.borrow().right, self.node.borrow().value.clone()));
        match &self.right_child {
            None => {},
            Some(right_child) => right_child.borrow().recurse_insert(v),
        };
    }
    fn recurse_find(&self, id: i32) -> Option<T> {
        return if id < self.node.borrow().right && id >= self.node.borrow().left {
            Some(self.node.borrow().value.clone())
        } else if id < self.node.borrow().left {
            match &self.left_child {
                None => None,
                Some(left_child) => left_child.borrow().recurse_find(id),
            }
        } else {
            match &self.right_child {
                None => None,
                Some(right_child) => right_child.borrow().recurse_find(id),
            }
        }
    }
}

impl<T: Clone + Default> InternalTree<T> {
    pub(crate) fn index(&self, id: i32) -> Option<T> {
        self.recurse_find(id)
    }

    pub(crate) fn update(&mut self, left: i32, right: i32, value: &T) -> Option<()> {
        if right - left < 1 {
            None
        } else {
            self.update_node(left, right, value);
            Some(())
        }
    }

    pub(crate) fn from_vec(internals: &Vec<(i32, i32, T)>) -> InternalTree<T> {
        let mut tree: InternalTree<T> = InternalTree::from_internal(internals[0].0, internals[0].1);
        internals.iter().for_each(|i| tree.update(i.0, i.1, &i.2).unwrap());
        tree
    }

    pub(crate) fn to_vec(self) -> Vec<(i32, i32, T)> {
        let mut v = vec![];
        self.recurse_insert(&mut v);
        return v;
    }
}

#[test]
fn test_internal() {
    let internals = vec![(4, 8, 4..8), (3, 10, 3..10), (5, 7, 5..7), (100, 200, 100..200)];
    let internal_tree = InternalTree::from_vec(&internals);
    println!("{:?}", internal_tree.index(6));
    println!("{:?}", internal_tree.index(7));
    println!("{:?}", internal_tree.index(60));
    println!("{:?}", internal_tree.index(100));
    println!("{:?}", internal_tree.index(99));
    println!("{:?}", internal_tree.index(200));
    let ordered_internals = internal_tree.to_vec();
    println!("{:?}", ordered_internals);
}