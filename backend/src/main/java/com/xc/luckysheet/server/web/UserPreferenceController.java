package com.xc.luckysheet.server.web;

import com.xc.luckysheet.server.contract.UserPreferences;
import com.xc.luckysheet.server.contract.UserPreferencesRequest;
import com.xc.luckysheet.server.service.ActorIdentity;
import com.xc.luckysheet.server.service.UserPreferenceService;
import jakarta.validation.Valid;
import org.springframework.security.core.Authentication;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PutMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/api/user-preferences")
public class UserPreferenceController {
    private final UserPreferenceService preferences;

    public UserPreferenceController(UserPreferenceService preferences) {
        this.preferences = preferences;
    }

    @GetMapping
    public UserPreferences get(Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return preferences.get(ActorIdentity.subject(authentication));
    }

    @PutMapping
    public UserPreferences update(@Valid @RequestBody UserPreferencesRequest request, Authentication authentication) {
        ActorIdentity.requireRegisteredActor(authentication);
        return preferences.update(ActorIdentity.subject(authentication), request);
    }
}
