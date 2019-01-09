import React from 'react'
import ReactDOM from 'react-dom'
import './index.css'
import registerServiceWorker from './registerServiceWorker'
import { IntlProvider, addLocaleData } from "react-intl"
import fr from "react-intl/locale-data/fr"
import en from "react-intl/locale-data/en"
import App from './App'
import $ from 'jquery'

(function () {

  addLocaleData([...en, ...fr])

  const Office = window.Office

  Office.initialize = function (reason) {
    $(document).ready(function () {

      const language = Office.context.displayLanguage.substring(0,2)
      const mailbox = Office.context.mailbox
      const email = mailbox._userProfile$p$0._data$p$0._data$p$0.userEmailAddress
      const settings = Office.context.roamingSettings

      const paramsOffice = {mailbox: mailbox, email: email, settings: settings}

      ReactDOM.render(
        <IntlProvider locale={language} defaultLocale="fr">
          <App paramsOffice={paramsOffice}/>
        </IntlProvider>
        , document.getElementById('root')
      )
      registerServiceWorker()
    });
  };
})();
