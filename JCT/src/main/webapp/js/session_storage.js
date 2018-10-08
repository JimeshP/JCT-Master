var ClientSession = {
    get: function (prm_key) {
        //Variable Declaration
        var _excpMsg;
        var _return;

        //Variable Declaration Block
        _return = "";
        _excpMsg = "Within ClientSession.get:\n";

        //Exception Handling Block
        try {
            /*
             * If prm_key Is Null, Throw Exception
             * Else, Return Values
             */
            if (prm_key == null) {
                throw "Key Not Provided";
            } else {
                //Get Value In _return
                _return = sessionStorage.getItem(prm_key);
            }
        } catch (_excp) {
            //Create Exception Message
            _excpMsg += _excp;
            //Display Message
            alert(_excpMsg);
        }

        //Return Statement
        return _return;
    },
    remove: function (prm_key) {
        //Variable Declaration
        var _excpMsg;

        //Variable Declaration Block
        _excpMsg = "Within ClientSession.remove:\n";

        //Exception Handling Block
        try {
            /*
             * If prm_key Is Null, Throw Exception
             * Else, Return Values
             */
            if (prm_key == null) {
                throw "Key Not Provided";
            } else {
                //Set Value In _return
                sessionStorage.removeItem(prm_key);
            }
        } catch (_excp) {
            //Create Exception Message
            _excpMsg += _excp;
            //Display Message
            alert(_excpMsg);
        }
    },
    removeAll: function () {
        //Variable Declaration
        var _excpMsg;

        //Variable Declaration Block
        _excpMsg = "Within ClientSession.removeAll:\n";

        //Exception Handling Block
        try {
            //Clear sessionStorage
            sessionStorage.clear();
        } catch (_excp) {
            //Create Exception Message
            _excpMsg += _excp;
            //Display Message
            alert(_excpMsg);
        }
    },
    set: function (prm_key, prm_val) {
        //Variable Declaration
        var _excpMsg;

        //Variable Declaration Block
        _excpMsg = "Within ClientSession.set:\n";

        //Exception Handling Block
        try {
            /*
             * If prm_key Is Null, Throw Exception
             * Else, Return Values
             */
            if (prm_key == null) {
                throw "Key Not Provided";
            } else {
                //Set Value In _return
                sessionStorage.setItem(prm_key, prm_val);
            }
        } catch (_excp) {
            //Create Exception Message
            _excpMsg += _excp;
            //Display Message
            alert(_excpMsg);
        }
    }
};