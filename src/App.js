import React, { useState, state } from "react";
import * as XLSX from "xlsx";
import "./App.css";
import { Base64 } from "js-base64";

class App extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      items: null,
      generated: false,
      htmlStr: null,
    };
  }
  render() {
    const readExcel = (file) => {
      const promise = new Promise((resolve, reject) => {
        const fileReader = new FileReader();
        fileReader.readAsArrayBuffer(file);

        fileReader.onload = (e) => {
          const bufferArray = e.target.result;
          const wb = XLSX.read(bufferArray, { type: "buffer" });
          const wsname = wb.SheetNames[0];
          const ws = wb.Sheets[wsname];
          const data = XLSX.utils.sheet_to_json(ws);
          resolve(data);
        };

        fileReader.onerror = (error) => {
          reject(error);
        };
      });

      let htmlStr = `[fusion_builder_container hundred_percent="no" hundred_percent_height="no"
      hundred_percent_height_scroll="no" hundred_percent_height_center_content="yes"
      equal_height_columns="no" menu_anchor=""
      hide_on_mobile="small-visibility,medium-visibility,large-visibility"
      status="published" publish_date="" class="" id="" border_size="" border_color=""
      border_style="solid" margin_top="50" margin_bottom="" padding_top=""
      padding_right="" padding_bottom="" padding_left="" gradient_start_color=""
      gradient_end_color="" gradient_start_position="0" gradient_end_position="100"
      gradient_type="linear" radial_direction="center" linear_angle="180"
      background_color="" background_image="" background_position="center center"
      background_repeat="no-repeat" fade="no" background_parallax="none"
      enable_mobile="no" parallax_speed="0.3" background_blend_mode="none"
      video_mp4="" video_webm="" video_ogv="" video_url="" video_aspect_ratio="16:9"
      video_loop="yes" video_mute="yes" video_preview_image="" filter_hue="0"
      filter_saturation="100" filter_brightness="100" filter_contrast="100"
      filter_invert="0" filter_sepia="0" filter_opacity="100" filter_blur="0"
      filter_hue_hover="0" filter_saturation_hover="100" filter_brightness_hover="100"
      filter_contrast_hover="100" filter_invert_hover="0" filter_sepia_hover="0"
      filter_opacity_hover="100"
      filter_blur_hover="0"][fusion_builder_row][fusion_builder_column type="1_1"
      layout="3_4" spacing="" center_content="no" link="" target="_self" min_height=""
      hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
      id="" hover_type="none" border_size="0" border_color="" border_style="solid"
      border_position="all" border_radius="" box_shadow="no" dimension_box_shadow=""
      box_shadow_blur="0" box_shadow_spread="0" box_shadow_color=""
      box_shadow_style="" padding_top="" padding_right="" padding_bottom=""
      padding_left="" margin_top="" margin_bottom="" background_type="single"
      gradient_start_color="" gradient_end_color="" gradient_start_position="0"
      gradient_end_position="100" gradient_type="linear" radial_direction="center"
      linear_angle="180" background_color="" background_image=""
      background_image_id="" background_position="left top"
      background_repeat="no-repeat" background_blend_mode="none" animation_type=""
      animation_direction="left" animation_speed="0.3" animation_offset=""
      filter_type="regular" filter_hue="0" filter_saturation="100"
      filter_brightness="100" filter_contrast="100" filter_invert="0" filter_sepia="0"
      filter_opacity="100" filter_blur="0" filter_hue_hover="0"
      filter_saturation_hover="100" filter_brightness_hover="100"
      filter_contrast_hover="100" filter_invert_hover="0" filter_sepia_hover="0"
      filter_opacity_hover="100" filter_blur_hover="0" last="no"][fusion_text
      columns="" column_min_width="" column_spacing="" rule_style="default"
      rule_size="" rule_color=""
      hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
      id="" animation_type="" animation_direction="left" animation_speed="0.3"
      animation_offset=""]
      <h3><strong>About</strong></h3>
      Attracting investment and successfully closing a funding round is a milestone
      for any startup. With the objective to foster startup ecosystems by increasing
      startups' visibility and attracting more attention from potential stakeholders,
      DealMatrix Deals Monitor provides an overview of the latest startup deal
      announcements in Europe and worldwide. The list is updated on a weekly basis. If
      we missed your deal feel free to send us a notification (see below).
      <strong>Last update:</strong> 01.03.2021
      <h3><strong>Key</strong></h3>
      Founded | HQ | Sector | Series/Stage Last Investment Raised | Valuation |
      Investors News Source
      [/fusion_text][/fusion_builder_column][fusion_builder_column type="1_1"
      layout="1_1" spacing="" center_content="no" link="" target="_self" min_height=""
      hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
      id="" hover_type="none" border_size="0" border_color="" border_style="solid"
      border_position="all" border_radius="" box_shadow="no" dimension_box_shadow=""
      box_shadow_blur="0" box_shadow_spread="0" box_shadow_color=""
      box_shadow_style="" padding_top="" padding_right="" padding_bottom=""
      padding_left="" margin_top="" margin_bottom="" background_type="single"
      gradient_start_color="" gradient_end_color="" gradient_start_position="0"
      gradient_end_position="100" gradient_type="linear" radial_direction="center"
      linear_angle="180" background_color="" background_image=""
      background_image_id="" background_position="left top"
      background_repeat="no-repeat" background_blend_mode="none" animation_type=""
      animation_direction="left" animation_speed="0.3" animation_offset=""
      filter_type="regular" filter_hue="0" filter_saturation="100"
      filter_brightness="100" filter_contrast="100" filter_invert="0" filter_sepia="0"
      filter_opacity="100" filter_blur="0" filter_hue_hover="0"
      filter_saturation_hover="100" filter_brightness_hover="100"
      filter_contrast_hover="100" filter_invert_hover="0" filter_sepia_hover="0"
      filter_opacity_hover="100" filter_blur_hover="0" last="no"][fusion_button
      link="#dealsubmission" text_transform="" title="" target="_self"
      link_attributes="" alignment="center" modal=""
      hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
      id="" color="default" button_gradient_top_color=""
      button_gradient_bottom_color="" button_gradient_top_color_hover=""
      button_gradient_bottom_color_hover="" accent_color="" accent_hover_color=""
      type="" bevel_color="" border_width="" border_radius="" border_color=""
      border_hover_color="" size="" stretch="yes" icon="" icon_position="left"
      icon_divider="no" animation_type="" animation_direction="left"
      animation_speed="0.3" animation_offset=""]I WANT MY DEAL ANNOUNCEMENT TO BE
      DISPLAYED ON THE LIST`;
      let htmlPanel;
      let companyURL;
      let companyName;
      let dealDate;
      let newsTitle;
      let foundedYear;
      let companySector;
      let companyStage;
      let dealInvestment;
      let investors;
      let valuation;
      let country;
      let link;
      promise.then((d) => {
        this.setState({ items: d }, () => {
          this.state.items.forEach((element) => {
            // console.log("yp", element["Company Name"]);

            companyURL = element["Company Website"];
            companyName = element["Company Name"];
            dealDate = element["Deal Announcement Date"];
            newsTitle = element["Deal News. Tag line/Title"];
            foundedYear = element["Company Founding Year"];
            companySector = element["Company Industry/Sector"];
            companyStage = element["Company Stage"];
            dealInvestment = element["Deal Investment Raised"];
            investors = element["Deal Investors"];
            valuation = element["Company Valuation"];
            country = element["Company Country"];
            link = element["Deal News Reference (link)"];
            htmlPanel = ` [/fusion_button][/fusion_builder_column][/fusion_builder_row][/fusion_builder_container][fusion_builder_container
              hundred_percent="no" hundred_percent_height="no"
              hundred_percent_height_scroll="no" hundred_percent_height_center_content="yes"
              equal_height_columns="no" menu_anchor=""
              hide_on_mobile="small-visibility,medium-visibility,large-visibility"
              status="published" publish_date="" class="" id="" border_size="" border_color=""
              border_style="solid" margin_top="" margin_bottom="" padding_top=""
              padding_right="" padding_bottom="" padding_left="" gradient_start_color=""
              gradient_end_color="" gradient_start_position="0" gradient_end_position="100"
              gradient_type="linear" radial_direction="center" linear_angle="180"
              background_color="" background_image="" background_position="center center"
              background_repeat="no-repeat" fade="no" background_parallax="none"
              enable_mobile="no" parallax_speed="0.3" background_blend_mode="none"
              video_mp4="" video_webm="" video_ogv="" video_url="" video_aspect_ratio="16:9"
              video_loop="yes" video_mute="yes" video_preview_image="" filter_hue="0"
              filter_saturation="100" filter_brightness="100" filter_contrast="100"
              filter_invert="0" filter_sepia="0" filter_opacity="100" filter_blur="0"
              filter_hue_hover="0" filter_saturation_hover="100" filter_brightness_hover="100"
              filter_contrast_hover="100" filter_invert_hover="0" filter_sepia_hover="0"
              filter_opacity_hover="100"
              filter_blur_hover="0"][fusion_builder_row][fusion_builder_column type="1_1"
              layout="1_1" spacing="" center_content="no" link="" target="_self" min_height=""
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" hover_type="none" border_size="0" border_color="" border_style="solid"
              border_position="all" border_radius="" box_shadow="no" dimension_box_shadow=""
              box_shadow_blur="0" box_shadow_spread="0" box_shadow_color=""
              box_shadow_style="" padding_top="" padding_right="" padding_bottom=""
              padding_left="" margin_top="" margin_bottom="" background_type="single"
              gradient_start_color="" gradient_end_color="" gradient_start_position="0"
              gradient_end_position="100" gradient_type="linear" radial_direction="center"
              linear_angle="180" background_color="" background_image=""
              background_image_id="" background_position="left top"
              background_repeat="no-repeat" background_blend_mode="none" animation_type=""
              animation_direction="left" animation_speed="0.3" animation_offset=""
              filter_type="regular" filter_hue="0" filter_saturation="100"
              filter_brightness="100" filter_contrast="100" filter_invert="0" filter_sepia="0"
              filter_opacity="100" filter_blur="0" filter_hue_hover="0"
              filter_saturation_hover="100" filter_brightness_hover="100"
              filter_contrast_hover="100" filter_invert_hover="0" filter_sepia_hover="0"
              filter_opacity_hover="100" filter_blur_hover="0" last="no"][fusion_separator
              style_type="none"
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" sep_color="" top_margin="30" bottom_margin="30" border_size="0" icon=""
              icon_circle="" icon_circle_color="" width="" alignment="center"
              /][/fusion_builder_column][/fusion_builder_row][/fusion_builder_container][fusion_builder_container
              hundred_percent="no" hundred_percent_height="no"
              hundred_percent_height_scroll="no" hundred_percent_height_center_content="yes"
              equal_height_columns="no" menu_anchor=""
              hide_on_mobile="small-visibility,medium-visibility,large-visibility"
              status="published" publish_date="" class="" id="" border_size="" border_color=""
              border_style="solid" margin_top="30" margin_bottom="" padding_top=""
              padding_right="30" padding_bottom="" padding_left="30" gradient_start_color=""
              gradient_end_color="" gradient_start_position="0" gradient_end_position="100"
              gradient_type="linear" radial_direction="center" linear_angle="180"
              background_color="#242424" background_image="" background_position="center
              center" background_repeat="no-repeat" fade="no" background_parallax="none"
              enable_mobile="no" parallax_speed="0.3" background_blend_mode="none"
              video_mp4="" video_webm="" video_ogv="" video_url="" video_aspect_ratio="16:9"
              video_loop="yes" video_mute="yes" video_preview_image="" filter_hue="0"
              filter_saturation="100" filter_brightness="100" filter_contrast="100"
              filter_invert="0" filter_sepia="0" filter_opacity="100" filter_blur="0"
              filter_hue_hover="0" filter_saturation_hover="100" filter_brightness_hover="100"
              filter_contrast_hover="100" filter_invert_hover="0" filter_sepia_hover="0"
              filter_opacity_hover="100"
              filter_blur_hover="0"][fusion_builder_row][fusion_builder_column type="1_1"
              layout="1_1" spacing="" center_content="no" link="" target="_self" min_height=""
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" hover_type="none" border_size="0" border_color="" border_style="solid"
              border_position="all" border_radius="" box_shadow="no" dimension_box_shadow=""
              box_shadow_blur="0" box_shadow_spread="0" box_shadow_color=""
              box_shadow_style="" padding_top="" padding_right="" padding_bottom=""
              padding_left="" margin_top="" margin_bottom="" background_type="single"
              gradient_start_color="" gradient_end_color="" gradient_start_position="0"
              gradient_end_position="100" gradient_type="linear" radial_direction="center"
              linear_angle="180" background_color="" background_image=""
              background_image_id="" background_position="left top"
              background_repeat="no-repeat" background_blend_mode="none" animation_type=""
              animation_direction="left" animation_speed="0.3" animation_offset=""
              filter_type="regular" filter_hue="0" filter_saturation="100"
              filter_brightness="100" filter_contrast="100" filter_invert="0" filter_sepia="0"
              filter_opacity="100" filter_blur="0" filter_hue_hover="0"
              filter_saturation_hover="100" filter_brightness_hover="100"
              filter_contrast_hover="100" filter_invert_hover="0" filter_sepia_hover="0"
              filter_opacity_hover="100" filter_blur_hover="0" last="no"][fusion_title
              title_type="text" rotation_effect="bounceIn" display_time="1200"
              highlight_effect="circle" loop_animation="off" highlight_width="9"
              highlight_top_margin="0" before_text="" rotation_text="" highlight_text=""
              after_text=""
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" content_align="left" size="1" font_size="" animated_font_size=""
              line_height="" letter_spacing="" margin_top="" margin_bottom=""
              margin_top_mobile="" margin_bottom_mobile="" text_color=""
              animated_text_color="" highlight_color="" style_type="default" sep_color=""]
              <h4>${dealDate}</h4>
              [/fusion_title][/fusion_builder_column][fusion_builder_column type="1_5"
              layout="1_5" spacing="" center_content="no" link="" target="_self" min_height=""
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" hover_type="none" border_size="0" border_color="" border_style="solid"
              border_position="all" border_radius="" box_shadow="no" dimension_box_shadow=""
              box_shadow_blur="0" box_shadow_spread="0" box_shadow_color=""
              box_shadow_style="" padding_top="" padding_right="" padding_bottom=""
              padding_left="" margin_top="" margin_bottom="" background_type="single"
              gradient_start_color="" gradient_end_color="" gradient_start_position="0"
              gradient_end_position="100" gradient_type="linear" radial_direction="center"
              linear_angle="180" background_color="" background_image=""
              background_image_id="" background_position="left top"
              background_repeat="no-repeat" background_blend_mode="none" animation_type=""
              animation_direction="left" animation_speed="0.3" animation_offset=""
              filter_type="regular" filter_hue="0" filter_saturation="100"
              filter_brightness="100" filter_contrast="100" filter_invert="0" filter_sepia="0"
              filter_opacity="100" filter_blur="0" filter_hue_hover="0"
              filter_saturation_hover="100" filter_brightness_hover="100"
              filter_contrast_hover="100" filter_invert_hover="0" filter_sepia_hover="0"
              filter_opacity_hover="100" filter_blur_hover="0" last="no"][fusion_separator
              style_type="none"
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" sep_color="" top_margin="30" bottom_margin="30" border_size="0" icon=""
              icon_circle="" icon_circle_color="" width="" alignment="center"
              /][fusion_imageframe image_id="7009|medium" max_width="250" style_type=""
              blur="" stylecolor="" hover_type="none" bordersize="" bordercolor=""
              borderradius="" align="none" lightbox="no" gallery_id="" lightbox_image=""
              lightbox_image_id="" alt="" link=${companyURL}
              linktarget="_blank"
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" animation_type="" animation_direction="left" animation_speed="0.3"
              animation_offset=""]https://dealmatrix.com/wp-content/uploads/2021/03/exscientia--300x300.png?_t=1614963320[/fusion_imageframe][/fusion_builder_column][fusion_builder_column
              type="4_5" layout="4_5" spacing="" center_content="no" link="" target="_self"
              min_height=""
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" hover_type="none" border_size="0" border_color="" border_style="solid"
              border_position="all" border_radius="" box_shadow="no" dimension_box_shadow=""
              box_shadow_blur="0" box_shadow_spread="0" box_shadow_color=""
              box_shadow_style="" padding_top="" padding_right="" padding_bottom=""
              padding_left="" margin_top="" margin_bottom="" background_type="single"
              gradient_start_color="" gradient_end_color="" gradient_start_position="0"
              gradient_end_position="100" gradient_type="linear" radial_direction="center"
              linear_angle="180" background_color="" background_image=""
              background_image_id="" background_position="left top"
              background_repeat="no-repeat" background_blend_mode="none" animation_type=""
              animation_direction="left" animation_speed="0.3" animation_offset=""
              filter_type="regular" filter_hue="0" filter_saturation="100"
              filter_brightness="100" filter_contrast="100" filter_invert="0" filter_sepia="0"
              filter_opacity="100" filter_blur="0" filter_hue_hover="0"
              filter_saturation_hover="100" filter_brightness_hover="100"
              filter_contrast_hover="100" filter_invert_hover="0" filter_sepia_hover="0"
              filter_opacity_hover="100" filter_blur_hover="0" last="no"][fusion_text
              columns="" column_min_width="" column_spacing="" rule_style="default"
              rule_size="" rule_color=""
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" animation_type="" animation_direction="left" animation_speed="0.3"
              animation_offset=""]
              <h3 style="text-align: left">
                <a href=${companyURL}>${companyName}</a> ${newsTitle}
              </h3>
              [/fusion_text][fusion_text columns="" column_min_width="" column_spacing=""
              rule_style="default" rule_size="" rule_color=""
              hide_on_mobile="small-visibility,medium-visibility,large-visibility" class=""
              id="" animation_type="" animation_direction="left" animation_speed="0.3"
              animation_offset=""]
              <p style="font-size: 18px">
                <strong
                  >Founded: ${foundedYear} | HQ: ${country} | Sector: 
                  ${companySector}|${companyStage}</strong
                >
              </p>
              <p style="font-size: 18px">
                <strong
                  >Investment Raised: ${dealInvestment} | Valuation: ${valuation} | Investors: </strong
                >${investors},
                <a href="https://www.bms.com/">Bristol-Myers Squibb</a>,
                <a href="https://www.evotec.com/en">Evotec</a>,
                <a href="http://www.gthcap.com/en/">GT Healthcare Capital Partners</a>,
                <a href="https://www.novoholdings.dk/">Novo Holdings</a>
              </p>
              <strong
                ><a
                  href="${link}"
                  >News Link</a
                ></strong
              >
            `;
            // console.log("yp", htmlPanel);
            htmlStr = htmlStr + htmlPanel;
            console.log("hey", htmlStr);
          });
          this.setState({ htmlStr: htmlStr, generated: true });
        });
      });
    };

    return (
      <div className="App">
        {this.state.generated ?  <h2>Click copy</h2> :  <h2>Upload xlsx file</h2>}
        <div className="App-container">
          <input
            type="file"
            onChange={(e) => {
              const file = e.target.files[0];
              readExcel(file);
            }}
          />
          {this.state.generated ? (
            <>
              <textarea className="text-area" value={this.state.htmlStr}/>
              <button
                onClick={() => {
                  navigator.clipboard.writeText(this.state.htmlStr)
                }}
              >
                Copy
              </button>
            </>
          ) : (
           ''
          )}
        </div>
      </div>
    );
  }
}
export default App;
