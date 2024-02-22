#[derive(Default)]
pub struct Properties<'a> {
    pub(crate) title: Option<&'a str>,
    pub(crate) subject: Option<&'a str>,
    pub(crate) author: Option<&'a str>,
    pub(crate) manager: Option<&'a str>,
    pub(crate) company: Option<&'a str>,
    pub(crate) category: Option<&'a str>,
    pub(crate) keywords: Option<&'a str>,
    pub(crate) comments: Option<&'a str>,
    pub(crate) status: Option<&'a str>,
}

impl<'a> Properties<'a> {
    pub fn set_title(&mut self, title: &'a str) -> &mut Self {
        self.title = Some(title);
        self
    }

    pub fn set_subject(&mut self, subject: &'a str) -> &mut Self {
        self.subject = Some(subject);
        self
    }

    pub fn set_author(&mut self, author: &'a str) -> &mut Self {
        self.author = Some(author);
        self
    }

    pub fn set_manager(&mut self, manager: &'a str) -> &mut Self {
        self.manager = Some(manager);
        self
    }

    pub fn set_company(&mut self, company: &'a str) -> &mut Self {
        self.company = Some(company);
        self
    }

    pub fn set_category(&mut self, category: &'a str) -> &mut Self {
        self.category = Some(category);
        self
    }

    pub fn set_keywords(&mut self, keywords: &'a str) -> &mut Self {
        self.keywords = Some(keywords);
        self
    }

    pub fn set_comments(&mut self, comments: &'a str) -> &mut Self {
        self.comments = Some(comments);
        self
    }

    pub fn set_status(&mut self, status: &'a str) -> &mut Self {
        self.status = Some(status);
        self
    }
}