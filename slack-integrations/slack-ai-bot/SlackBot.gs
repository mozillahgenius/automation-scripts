class SlackBot {
    constructor(e) {
      this.requestEvent = e;
      this.postData = null;
      this.slackEvent = null;
      this.responseData = this.init();
      this.verification();
    }
  
    responseJsonData(json) {
      return ContentService.createTextOutput(JSON.stringify(json)).setMimeType(ContentService.MimeType.JSON);
    }
  
    init() {
      const e = this.requestEvent;
      if (!e?.postData) return { error: 'postData is missing or undefined.', request: JSON.stringify(e, null, "  ") };
      this.postData = e.postData;
      if (!this.postData?.type) return { error: 'postData type is missing or undefined.', request: JSON.stringify(this.postData, null, "  ") };
      try { var event = JSON.parse(this.postData.contents); }
      catch (error) {
        event = e.parameter?.command && e.parameter?.text ? { event: { type: "command", event: { ...e.parameter } } } : { error: 'Invalid JSON format in postData contents.', request: this.postData };
      }
      this.slackEvent = event;
      return event?.event ? null : { error: 'Slack event is missing or undefined.', request: JSON.stringify(e, null, "  ") };
    }
  
    verification() {
      //SheetLog.log(JSON.stringify(this.postData));
      if (!this.postData || this.responseData) return null;
      if (this.postData.type !== 'url_verification') return null;
      this.responseData = { "challenge": this.postData.challenge };
      return this.responseData;
    }
  
    hasCache(key) {
      if (!key) return true;
      const cache = CacheService.getScriptCache();
      const cached = cache.get(key);
      if (cached) return true;
      cache.put(key, true, 30 * 60);
      return false;
    }
  
    handleEvent(type, callback = () => { }) {
      if (!this.slackEvent || this.responseData || this.slackEvent?.event?.type !== type) return null;
      const callbackResponse = callback({ event: this.slackEvent.event });
      if (!callbackResponse) return null;
      this.responseData = callbackResponse;
      return callbackResponse;
    }
  
    handleBase(type, targetType, callback = () => {}) {
      return this.handleEvent(type, ({ event }) => {
        const { text: message, channel, thread_ts: threadTs, ts, client_msg_id, bot_id, app_id } = event;
        if (bot_id || app_id) return null;
        if (event.type !== targetType || this.hasCache(`${channel}:${client_msg_id}`)) return null;
        return callback ? callback({ message, channel, threadTs: threadTs ?? ts, event }) : null;
      });
    }
  
    handleMessageEventBase(callback) { 
      return this.handleBase("message", "message", callback); }
    handleMentionEventBase(callback) { 
      return this.handleBase("app_mention", "app_mention", callback); }
    handleReactionEventBase(callback) { 
      return this.handleBase("reaction_added", "reaction_added", callback); }
  
    response() {
      Logger.log(this.responseData);
      return this.responseData && this.responseJsonData(this.responseData);
    }
  }
  