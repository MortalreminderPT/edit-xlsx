pub struct Filters<'a> {
    pub(crate) and: Option<u8>,
    pub(crate) blank: Option<u8>,
    pub(crate) filters: Vec<Filter<'a>>,
}

impl<'a> Filters<'a> {
    pub fn new() -> Self {
        Self {
            and: None,
            blank: None,
            filters: vec![],
        }
    }
    pub fn eq(vals: Vec<&'a str>) -> Self {
        Self {
            and: None,
            blank: None,
            filters: vals.iter().map(|&v| Filter::eq(v)).collect(),
        }
    }
    pub fn blank() -> Self {
        Self {
            and: None,
            blank: Some(1),
            filters: vec![],
        }
    }
    pub fn not_blank() -> Self {
        Self {
            and: None,
            blank: None,
            filters: vec![Filter::ne(" ")],
        }
    }
    pub fn or(&mut self, filter: Filter<'a>) -> &mut Self {
        self.filters.push(filter);
        self
    }
    pub fn and(&mut self, filter: Filter<'a>) -> &mut Self {
        self.and = Some(1);
        self.filters.push(filter);
        self
    }

    pub(crate) fn is_custom_filters(&self) -> bool {
        self.filters.iter().any(|f| f.custom_filter)
    }
}

pub struct Filter<'a> {
    pub(crate) val: &'a str,
    pub(crate) operator: Option<&'a str>,
    custom_filter: bool,
}

impl<'a> Filter<'a> {
    pub fn eq(val: &'a str) -> Self {
        Self {
            val,
            operator: None,
            custom_filter: false,
        }
    }
    pub fn gt(val: &'a str) -> Self {
        Self {
            val,
            operator: Some("greaterThan"),
            custom_filter: true,
        }
    }
    pub fn lt(val: &'a str) -> Self {
        Self {
            val,
            operator: Some("lessThan"),
            custom_filter: true,
        }
    }

    pub fn ne(val: &'a str) -> Self {
        Self {
            val,
            operator: Some("notEqual"),
            custom_filter: true,
        }
    }
}