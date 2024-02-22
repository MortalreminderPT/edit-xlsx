pub struct Filters<'a> {
    pub(crate) and: u8,
    pub(crate) filters: Vec<Filter<'a>>,
}

impl<'a> Filters<'a> {
    pub fn new() -> Self {
        Self {
            and: 0,
            filters: vec![],
        }
    }

    pub fn or(&mut self, filter: Filter<'a>) -> &mut Self {
        self.and = 0;
        self.filters.push(filter);
        self
    }

    pub fn and(&mut self, filter: Filter<'a>) -> &mut Self {
        self.and = 1;
        self.filters.push(filter);
        self
    }
}

pub struct Filter<'a> {
    pub(crate) val: &'a str,
    pub(crate) operator: Option<&'a str>,
}

impl<'a> Filter<'a> {
    pub fn eq(val: &'a str) -> Self {
        Self {
            val,
            operator: None,
        }
    }
}